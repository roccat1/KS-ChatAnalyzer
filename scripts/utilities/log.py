import datetime

import scripts.gvar as gvar


def log(msg: str) -> None:
    """Logs a message to the log file

    Args:
        msg (str): Message to log
    """
    print(msg)

    with open(gvar.logPath, "a", encoding="utf8") as f:
        f.write(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")+" - "+msg+"\n")
