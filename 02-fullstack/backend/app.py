from pathlib import Path
from uuid import uuid4
import asyncio
import re
import sys
import zipfile
import tempfile
import shutil
import traceback

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

# O código do projeto pode estar em OneDrive/rede, mas os arquivos de uma
# execução ficam em uma área temporária LOCAL do Windows. Isso evita erros
# de caminho/sincronização durante a geração do Excel.
ROOT = Path(__file__).resolve().parents[1]
CORE = ROOT / "core"
FRONTEND = ROOT / "frontend"
LOCAL_APP_DIR = Path(tempfile.gettempdir()) / "RH_Valoracao_FullStack"
RUNS = LOCAL_APP_DIR / "runs"
RUNS.mkdir(parents=True, exist_ok=True)
sys.path.insert(0, str(CORE))
import rh_valoracao_core as core

MAX_TOTAL_UPLOAD = 150 * 1024 * 1024
ALLOWED = {".xlsx", ".xlsm", ".pdf", ".zip"}
RUN_ID_RE = re.compile(r"^[0-9a-f]{32}$")

app = FastAPI(
    title="RH Valoração — Processador Universal",
    version="4.1",
    description="Aplicação local para processar a base RH Valoração com folhas de pagamento em PDF.",
)
lock = asyncio.Lock()
app.mount("/static", StaticFiles(directory=FRONTEND), name="static")


@app.get("/")
def index():
    return FileResponse(FRONTEND / "index.html")


@app.get("/health")
def health():
    return {"ok": True, "service": "rh-valoracao", "version": "4.1"}


def _safe_relative_name(name: str) -> Path:
    raw = (name or "arquivo").replace("\\", "/")
    parts = [p for p in raw.split("/") if p not in ("", ".", "..")]
    if not parts:
        return Path("arquivo")
    return Path(*parts)


def _validate_upload_name(name: str):
    suffix = Path(name or "").suffix.lower()
    if suffix not in ALLOWED:
        raise HTTPException(400, detail=f"Formato não aceito: {Path(name or 'arquivo').name}. Use Excel, PDF ou ZIP.")


def _resumo_legivel(resumo):
    out = []
    for item in resumo:
        out.append({
            "competencia": item.get("competencia", ""),
            "formato": item.get("formato", ""),
            "arquivo": Path(item.get("arquivo", "")).name,
            "colaboradores_modelo": item.get("colaboradores_modelo", 0),
            "encontrados": item.get("encontrados", 0),
            "nao_encontrados": item.get("nao_encontrados", 0),
            "lancamentos": item.get("lancamentos", 0),
            "valor_total_extraido": round(float(item.get("valor_total_extraido", 0) or 0), 2),
        })
    return out


def _erro_publico(exc: Exception) -> str:
    msg = str(exc).strip()
    if not msg:
        return f"O processamento encontrou um erro inesperado ({type(exc).__name__})."
    return f"{type(exc).__name__}: {msg[:1200]}"


def _registrar_erro(run_dir: Path, exc: Exception):
    try:
        log = run_dir / "ERRO_EXECUCAO.txt"
        log.write_text(traceback.format_exc(), encoding="utf-8")
    except Exception:
        pass


@app.post("/processar")
async def processar(files: list[UploadFile] = File(...)):
    async with lock:
        run_id = uuid4().hex
        run_dir = RUNS / run_id
        run_dir.mkdir(parents=True, exist_ok=True)
        total = 0
        try:
            if not files:
                raise HTTPException(400, detail="Nenhum arquivo recebido.")

            for item in files:
                nome = item.filename or "arquivo"
                _validate_upload_name(nome)
                rel = _safe_relative_name(nome)
                dest = run_dir / rel
                dest.parent.mkdir(parents=True, exist_ok=True)
                data = await item.read()
                total += len(data)
                if total > MAX_TOTAL_UPLOAD:
                    raise HTTPException(413, detail="Os arquivos enviados ultrapassam o limite de 150 MB por execução.")
                if not data:
                    raise HTTPException(400, detail=f"Arquivo vazio: {Path(nome).name}.")
                dest.write_bytes(data)

            # O motor recebe uma pasta LOCAL, fora do diretório do projeto.
            # Isso é importante quando o projeto estiver dentro do OneDrive.
            core.PASTA_TRABALHO = str(run_dir)
            if hasattr(core, "_ZIPS_EXTRAIDOS"):
                core._ZIPS_EXTRAIDOS.clear()

            excels, pdfs = core.localizar_arquivos()
            if len(excels) != 1:
                raise HTTPException(400, detail=f"É necessário exatamente 1 Excel-base (.xlsx/.xlsm). Encontrados: {len(excels)}.")
            if not pdfs:
                raise HTTPException(400, detail="Nenhum PDF encontrado, inclusive após a extração dos ZIPs.")

            output, relatorio = core.processar()
            output = Path(output).resolve()
            if not output.exists() or output.stat().st_size == 0:
                raise RuntimeError("O processamento terminou sem gerar um arquivo Excel válido.")

            resumo = _resumo_legivel(core.gerar_relatorio_execucao(relatorio))
            return {
                "ok": True,
                "run_id": run_id,
                "arquivo": output.name,
                "download_url": f"/download/{run_id}",
                "entrada": {
                    "excel": [Path(x).name for x in excels],
                    "pdfs": [Path(x).name for x in pdfs],
                    "quantidade_pdf": len(pdfs),
                },
                "resumo": resumo,
            }
        except HTTPException:
            raise
        except (ValueError, zipfile.BadZipFile) as exc:
            _registrar_erro(run_dir, exc)
            raise HTTPException(400, detail=_erro_publico(exc))
        except Exception as exc:
            _registrar_erro(run_dir, exc)
            raise HTTPException(500, detail=_erro_publico(exc))


@app.get("/download/{run_id}")
def download(run_id: str):
    if not RUN_ID_RE.fullmatch(run_id or ""):
        raise HTTPException(400, detail="Execução inválida.")
    run_dir = RUNS / run_id
    if not run_dir.exists():
        raise HTTPException(404, detail="Execução não encontrada.")
    outputs = [p for p in run_dir.rglob("*_V11_ATUALIZADA.*") if p.suffix.lower() in {".xlsx", ".xlsm"}]
    if not outputs:
        raise HTTPException(404, detail="Arquivo final não encontrado para esta execução.")
    output = outputs[0].resolve()
    return FileResponse(output, filename=output.name)
