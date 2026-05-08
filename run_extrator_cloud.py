"""
Runner cloud para o Extrator G.Z.U.S.

Versão ajustada para GitHub Actions:
- carrega o extrator .pyw sem abrir Tkinter;
- força Selenium em modo headless;
- aumenta janela/zoom para reduzir diferença entre PC local e nuvem;
- reforça a etapa de exportação, que no headless pode demorar ou ficar fora da área visível;
- roda processamento e deixa o extrator gerar/enviar o SQLite.
"""
from __future__ import annotations

import importlib.machinery
import importlib.util
import os
import sys
import time
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

        # No GitHub Actions o Chrome headless pode abrir com viewport diferente do PC.
        # Forçamos uma área grande e um zoom menor para o botão "Ações" não sumir.
        try:
            self.driver.set_window_size(1920, 1200)
            self.driver.execute_script("document.body.style.zoom='80%'")
            log("🖥️ Viewport cloud ajustado para 1920x1200 com zoom 80%.")
        except Exception as e:
            log(f"⚠️ Não consegui ajustar viewport/zoom: {e}")

        # O wait original é 25s. Na nuvem algumas telas demoram mais.
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

        # Garante que nenhum menu/popup esteja bloqueando a tela.
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

        # Reforça viewport/zoom também no momento da exportação.
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
            log("⚠️ Não encontrei o botão Ações com os seletores reforçados. Tentando exportação original...")
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
