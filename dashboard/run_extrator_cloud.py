"""
Runner cloud para o Extrator G.Z.U.S.

O que este arquivo faz:
- carrega o extrator .pyw sem abrir a tela Tkinter;
- força Selenium em modo headless;
- roda a extração;
- roda o processamento;
- deixa o próprio extrator gerar o SQLite leve e enviar para o GitHub.

Este arquivo deve ficar na raiz do repositório painel-faturamento,
junto do app.py, banco_gzus.py e do arquivo .pyw do extrator.
"""
from __future__ import annotations

import importlib.machinery
import importlib.util
import os
import sys
import traceback
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent

EXTRATOR_ARQUIVO = os.getenv(
    "EXTRATOR_ARQUIVO",
    "Extrator_19_EAs_SAO_MIGUEL_TRATAMENTO_CORRIGIDO_GITHUB_502_CORRIGIDO.pyw",
)


def log(msg: str) -> None:
    print(msg, flush=True)


def importar_extrator(caminho: Path):
    """Importa .py ou .pyw de forma compatível com Linux/GitHub Actions."""
    if not caminho.exists():
        candidatos = sorted(BASE_DIR.glob("Extrator_19_EAs*.pyw")) + sorted(BASE_DIR.glob("Extrator_19_EAs*.py"))
        if candidatos:
            caminho = candidatos[-1]
        else:
            raise FileNotFoundError(f"Extrator não encontrado: {caminho}")

    loader = importlib.machinery.SourceFileLoader("extrator_gzus", str(caminho))
    spec = importlib.util.spec_from_loader("extrator_gzus", loader)
    if spec is None:
        raise RuntimeError(f"Não consegui preparar a importação do extrator: {caminho}")

    modulo = importlib.util.module_from_spec(spec)
    sys.modules["extrator_gzus"] = modulo
    loader.exec_module(modulo)
    return modulo, caminho


def main() -> int:
    load_dotenv(BASE_DIR / ".env", override=True)

    # Execução em nuvem, sem tela.
    os.environ["HEADLESS"] = "1"
    os.environ.setdefault("GZUS_POS_EXTRATOR_AUTO", "true")
    os.environ.setdefault("GITHUB_DASHBOARD_REMOTE_DIR", "dashboard")
    os.environ.setdefault("GZUS_PROJETO_PAINEL_DIR", str(BASE_DIR))

    caminho = BASE_DIR / EXTRATOR_ARQUIVO
    modulo, caminho_usado = importar_extrator(caminho)
    log(f"🚀 Extrator carregado: {caminho_usado.name}")

    eas_env = os.getenv("EAS_SELECIONADAS", "").strip()
    if eas_env:
        eas = [e.strip() for e in eas_env.split(";") if e.strip()]
    else:
        eas = list(getattr(modulo, "DEFAULT_EAS"))

    log(f"📌 EAs selecionadas: {len(eas)}")
    log("🌐 Iniciando Selenium em modo headless...")

    bot = modulo.ExtratorProducao(fila=None, headless=True, eas=eas)
    bot.executar()
    log("✅ Extração concluída. Iniciando processamento...")

    processamento = modulo.processar_arquivos_baixados(eas_list=eas, logger=log)
    log("✅ Processamento concluído.")

    if isinstance(processamento, dict):
        log(f"📊 Linhas lidas: {processamento.get('linhas_lidas', 0)}")
        log(f"📊 Linhas finais: {processamento.get('linhas_finais', 0)}")
        log(f"📁 Dashboard: {processamento.get('dashboard', '')}")

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print("❌ Falha na execução cloud:", exc, flush=True)
        traceback.print_exc()
        raise SystemExit(1)
