# CORREÇÃO 1 — colocar no começo do run_extrator_cloud.py, logo depois dos imports:

os.environ.setdefault("TZ", "America/Sao_Paulo")
try:
    time.tzset()
except Exception:
    pass


# CORREÇÃO 2 — dentro de main(), logo depois de:
# os.environ["HEADLESS"] = "1"
# adicionar:

os.environ.setdefault("TZ", "America/Sao_Paulo")
try:
    time.tzset()
except Exception:
    pass

log(f"🕒 Fuso horário do runner: {os.environ.get('TZ')} | agora local: {time.strftime('%d/%m/%Y %H:%M:%S')}")


# CORREÇÃO 3 — trocar este bloco:
#
# processamento = modulo.processar_arquivos_baixados(eas_list=eas, logger=log)
# log("✅ Processamento concluído.")
#
# por este bloco:

try:
    processamento = modulo.processar_arquivos_baixados(eas_list=eas, logger=log)
except Exception as exc:
    texto_erro = f"{type(exc).__name__}: {exc}".lower()
    erro_sem_csv = (
        "nenhum csv encontrado" in texto_erro
        or "nenhum arquivo encontrado" in texto_erro
        or "pasta downloads" in texto_erro
        or "período selecionado" in texto_erro
        or "periodo selecionado" in texto_erro
    )

    if erro_sem_csv:
        log("⚠️ Nenhum CSV novo foi encontrado para processar.")
        log("✅ Execução encerrada sem erro para preservar o dashboard/histórico anterior.")
        log("ℹ️ Isso evita falha do cron quando não há arquivo exportado ou quando o dia ainda não tem dados.")
        return 0

    raise

log("✅ Processamento concluído.")
