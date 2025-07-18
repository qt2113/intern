import logging
import colorlog
from logging.handlers import TimedRotatingFileHandler


log_colors_config = {
    'DEBUG': 'white',  # cyan white
    'INFO': 'green',
    'WARNING': 'yellow',
    'ERROR': 'red',
    'CRITICAL': 'bold_red',
}

def mylogger(filename, name='logger', file_level='INFO'):
    logger = logging.getLogger(name)
    # logger.setLevel(level=logging.DEBUG)

    # 输出到控制台
    console_handler = logging.StreamHandler()
    # 输出到文件
    file_handler = logging.FileHandler(filename=filename, mode='a', encoding='utf8')

    # 日志级别，logger 和 handler以最高级别为准，不同handler之间可以不一样，不相互影响
    logger.setLevel(logging.DEBUG)
    console_handler.setLevel(logging.DEBUG)
    file_handler.setLevel(eval('logging.' + file_level))    # 默认记录 INFO 以上级别日志到文件中

    # 日志输出格式
    file_formatter = logging.Formatter(
        fmt='%(asctime)s - %(levelname)s: %(message)s',
    )
    console_formatter = colorlog.ColoredFormatter(
        fmt='%(log_color)s%(asctime)s - %(levelname)s: %(message)s',
        # datefmt='%Y-%m-%d %H:%M:%S',
        log_colors=log_colors_config
    )
    # file_handler = TimedRotatingFileHandler(filename, when="D", interval=1, backupCount=10, encoding="UTF-8", utc=True)

    console_handler.setFormatter(console_formatter)
    file_handler.setFormatter(file_formatter)

    if not logger.handlers:
        logger.addHandler(console_handler)
        logger.addHandler(file_handler)

    console_handler.close()
    file_handler.close()

    return logger


# if __name__ == '__main__':
#     logger = mylogger('test.log')
#
#     logger.debug('debug')
#     logger.info('info')
#     logger.warning('warning')
#     logger.error('error')
#     logger.critical('critical')