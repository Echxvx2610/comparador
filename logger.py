import logging

def setup_logger(log_file):
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    # Crear un formateador de registro
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    # Crear un manejador para escribir los registros en un archivo
    file_handler = logging.FileHandler(log_file)
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    # Crear un manejador para imprimir los registros en la consola
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    # Agregar los manejadores al logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger
