import logger

# Configurar el logger y obtener una instancia del logger
logger = logger.setup_logger(r'comparador\app.log')

def funcion_ejemplo():
    logger.debug('Este es un mensaje de depuración')
    logger.info('Esta es una información')
    logger.warning('Este es un mensaje de advertencia')
    logger.error('Este es un mensaje de error')
    logger.critical('Este es un mensaje de critico')
    logger.debug("*"*40)

funcion_ejemplo()
