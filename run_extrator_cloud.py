"""
Runner cloud para o Extrator G.Z.U.S.

Versão cloud segura:
- força fuso horário America/Sao_Paulo no GitHub Actions;
- carrega o extrator .pyw sem abrir Tkinter;
- força Selenium headless;
- antes de extrair, copia o histórico já existente em dashboard/*.csv do repositório
  para output/dashboard, que é onde o extrator procura o histórico acumulado;
- mantém os ajustes de tela/zoom para exportação no GitHub Actions;
- roda processamento, recria gzus_dashboard.db e envia para o GitHub;
- se não houver CSV novo para processar, não derruba o workflow: mantém o painel anterior.
"""

from __future__ import annotations

import importlib.machinery
import importlib.util
import os
import shutil
import sys
import time
import traceback
from pathlib import Path

from dotenv import load_dotenv

# =========================
# FUSO HORÁRIO BRASIL
# =========================
os.environ.setdefault("TZ", "America/Sao_Paulo")
try:
    time.tzset()
except Exception:
    pass

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


def semear_historico_dashboard_no_output(modulo) -> None:
    """Copia dashboard/*.csv do repositório para output/dashboard.

    Essencial para o GitHub Actions:
    - a máquina nasce limpa;
    - o histórico confiável está versionado em dashboard/*.csv;
    - o extrator/processamento espera o histórico em output/dashboard.
    """
    repo_dashboard = BASE_DIR / "dashboard"

    process_output = getattr(modulo, "PROCESS_OUTPUT_FOLDER", None)
    if process_output:
        output_dashboard = Path(process_output) / "dashboard"
    else:
        output_dashboard = BASE_DIR / "output" / "dashboard"

    output_dashboard.mkdir(parents=True, exist_ok=True)

    arquivos_criticos = [
        "notas_dashboard.csv",
        "faturamento_contratos_dashboard.csv",
        "faturamento_dias_dashboard.csv",
        "faturamento_carro_estimado_dashboard.csv",
        "faturamento_carro_dias_dashboard.csv",
    ]

    copiados = 0
    for nome in arquivos_criticos:
        origem = repo_dashboard / nome
        destino = output_dashboard / nome
        if origem.exists():
            shutil.copy2(origem, destino)
            copiados += 1
            log(f"🌱 Histórico semeado no output/dashboard: {nome} ({origem.stat().st_size/1024:,.1f} KB)")
        else:
            log(f"⚠️ Histórico não encontrado no repositório: dashboard/{nome}")

    if copiados == 0:
        log("🚨 ATENÇÃO: nenhum CSV histórico foi semeado. O dashboard pode ficar só com a extração atual.")
    else:
        log(f"✅ Histórico cloud preparado: {copiados} arquivo(s) copiado(s) para {output_dashboard}")


def aplicar_ajustes_cloud(modulo):
    """Aplica pequenos ajustes sem alterar o extrator original."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait

    classe = modulo.ExtratorProducao
    preparar_original = classe.preparar
    exportar_original = classe.exportar

    def preparar_cloud(self):
        preparar_original(self)

        try:
            self.driver.set_window_size(1920, 1200)
            self.driver.execute_script("document.body.style.zoom='80%'")
            log("🖥️ Viewport cloud ajustado para 1920x1200 com zoom 80%.")
        except Exception as e:
            log(f"⚠️ Não consegui ajustar viewport/zoom: {e}")

        try:
            self.wait = WebDriverWait(self.driver, 60)
            log("⏱️ Timeout Selenium ampliado para 60s.")
        except Exception:
            pass

    def _salvar_debug(self, prefixo="debug_exportar"):
        try:
            debug_dir = BASE_DIR / "debug_github_actions"
            debug_dir.mkdir(exist_ok=True)
            png = debug_dir / f"{prefixo}.png"
            html = debug_dir / f"{prefixo}.html"
            self.driver.save_screenshot(str(png))
            html.write_text(self.driver.page_source or "", encoding="utf-8", errors="ignore")
            log(f"🧪 Debug salvo: {png}")
            log(f"🧪 HTML salvo: {html}")
        except Exception as e:
            log(f"⚠️ Não consegui salvar debug: {e}")

    def exportar_cloud(self):
        log("💾 Exportando...")

        try:
            self.fechar_menu_exibir_se_aberto()
            time.sleep(0.6)
        except Exception:
            pass

        try:
            body = self.driver.find_element(By.TAG_NAME, "body")
            body.send_keys(Keys.ESCAPE)
            time.sleep(0.4)
        except Exception:
            pass

        try:
            self.driver.set_window_size(1920, 1200)
            self.driver.execute_script("document.body.style.zoom='80%'")
            self.driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(0.5)
        except Exception:
            pass

        wait_longo = WebDriverWait(self.driver, 75)

        seletores_acoes = [
            (By.CSS_SELECTOR, "[aria-label='Ações']"),
            (By.CSS_SELECTOR, "[title='Ações']"),
            (By.XPATH, "//*[@aria-label='Ações' or @title='Ações']"),
            (By.XPATH, "//*[contains(normalize-space(.), 'Ações') and (self::button or @role='button' or ancestor::button)]"),
        ]

        ultimo_erro = None
        acoes_btn = None

        for by, seletor in seletores_acoes:
            try:
                log(f"🔎 Procurando botão Ações por: {seletor}")
                acoes_btn = wait_longo.until(EC.presence_of_element_located((by, seletor)))
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", acoes_btn)
                time.sleep(0.4)
                wait_longo.until(EC.element_to_be_clickable((by, seletor)))
                break
            except Exception as e:
                ultimo_erro = e
                acoes_btn = None

        if acoes_btn is None:
            log("⚠️ Não encontrei o botão Ações com seletores reforçados. Tentando exportação original...")
            try:
                return exportar_original(self)
            except Exception:
                _salvar_debug(self, "erro_sem_botao_acoes")
                raise ultimo_erro

        try:
            self.click_safe(acoes_btn)
        except Exception:
            try:
                self.driver.execute_script("arguments[0].click();", acoes_btn)
            except Exception as e:
                _salvar_debug(self, "erro_click_acoes")
                raise e

        time.sleep(1.2)

        seletores_exportar = [
            (By.XPATH, "//span[contains(normalize-space(.), 'Exportar')]"),
            (By.XPATH, "//*[contains(normalize-space(.), 'Exportar')]"),
            (By.XPATH, "//*[@role='menuitem' and contains(normalize-space(.), 'Exportar')]"),
        ]

        exportar_el = None
        ultimo_erro = None

        for by, seletor in seletores_exportar:
            try:
                log(f"🔎 Procurando opção Exportar por: {seletor}")
                exportar_el = wait_longo.until(EC.presence_of_element_located((by, seletor)))
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", exportar_el)
                time.sleep(0.3)
                break
            except Exception as e:
                ultimo_erro = e
                exportar_el = None

        if exportar_el is None:
            _salvar_debug(self, "erro_sem_opcao_exportar")
            raise ultimo_erro

        try:
            self.click_safe(exportar_el)
        except Exception:
            self.driver.execute_script("arguments[0].click();", exportar_el)

        log("📦 Exportado!")
        time.sleep(3)

    classe.preparar = preparar_cloud
    classe.exportar = exportar_cloud
    log("🛠️ Ajustes cloud aplicados ao Selenium/exportação.")


def _eh_erro_sem_csv(exc: Exception) -> bool:
    texto = f"{type(exc).__name__}: {exc}".lower()
    sinais = [
        "nenhum csv encontrado",
        "nenhum arquivo encontrado",
        "pasta downloads",
        "período selecionado",
        "periodo selecionado",
    ]
    return any(s in texto for s in sinais)


def main() -> int:
    load_dotenv(BASE_DIR / ".env", override=True)

    os.environ["HEADLESS"] = "1"
    os.environ.setdefault("TZ", "America/Sao_Paulo")
    try:
        time.tzset()
    except Exception:
        pass

    os.environ.setdefault("GZUS_POS_EXTRATOR_AUTO", "true")
    os.environ.setdefault("GITHUB_DASHBOARD_REMOTE_DIR", "dashboard")
    os.environ.setdefault("GZUS_PROJETO_PAINEL_DIR", str(BASE_DIR))

    caminho = BASE_DIR / EXTRATOR_ARQUIVO
    modulo, caminho_usado = importar_extrator(caminho)
    log(f"🚀 Extrator carregado: {caminho_usado.name}")
    log(f"🕒 Fuso horário do runner: {os.environ.get('TZ')} | agora local: {time.strftime('%d/%m/%Y %H:%M:%S')}")

    semear_historico_dashboard_no_output(modulo)
    aplicar_ajustes_cloud(modulo)

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

    try:
        processamento = modulo.processar_arquivos_baixados(eas_list=eas, logger=log)
    except Exception as exc:
        if _eh_erro_sem_csv(exc):
            log("⚠️ Nenhum CSV novo foi encontrado para processar.")
            log("✅ Execução encerrada sem erro para preservar o dashboard/histórico anterior.")
            log("ℹ️ Isso evita falha do cron quando não há arquivo exportado ou quando o dia ainda não tem dados.")
            return 0
        raise

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
