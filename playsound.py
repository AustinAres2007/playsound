import asyncio
import logging
from typing import Any

logger = logging.getLogger(__name__)


class PlaysoundException(Exception):
    pass

def winCommand(*command):

    '''
        Utilizes windll.winmm. Tested and known to work with MP3 and WAVE on
        Windows 7 with Python 2.7. Probably works with more file formats.
        Probably works on Windows XP thru Windows 10. Probably works with all
        versions of Python.

        Inspired by (but not copied from) Michael Gundlach <gundlach@gmail.com>'s mp3play:
        https://github.com/michaelgundlach/mp3play

        I never would have tried using windll.winmm without seeing his code.
    '''

    from ctypes import create_unicode_buffer, windll, wintypes

    windll.winmm.mciSendStringW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.UINT, wintypes.HANDLE]
    windll.winmm.mciGetErrorStringW.argtypes = [wintypes.DWORD, wintypes.LPWSTR, wintypes.UINT]

    bufLen = 600
    buf = create_unicode_buffer(bufLen)
    command = ' '.join(command)
    errorCode = int(windll.winmm.mciSendStringW(command, buf, bufLen - 1, 0))  # use widestring version of the function

    if errorCode:
        errorBuffer = create_unicode_buffer(bufLen)
        windll.winmm.mciGetErrorStringW(errorCode, errorBuffer, bufLen - 1)  # use widestring version of the function

        raise PlaysoundException("Error Code: {}, {}".format(errorCode, errorBuffer.value))

    return buf.value

def pause():
    try:
        winCommand(u'pause media')
    except PlaysoundException as err:
        errCode = int(str(err.args[0])[12:15])
        if errCode == 263:
            logger.error("Cannot pause, are you playing any media?")

def resume():
    try:
        winCommand(u'resume media')
    except PlaysoundException as err:
        errCode = int(str(err.args[0])[12:15])
        if errCode == 263:
            logger.error("Cannot resume, are you playing any media?")

'''
    Sends raw input to MCI (Can be anything, but will give an error if the syntax is incorrect)
'''
def send_raw(*command) -> Any:
    command = ' '.join(command)
    print(command)
    responce = winCommand(u'{}'.format(command))

    return responce

'''
    Seeks to the time specified.
'''
def seek(seconds: int):
    winCommand(u'seek media to {}'.format(seconds*1000))
    winCommand(u'play media')

async def play(sound):

    # All changes here are not done, I just need to winCommand(u'close media') after the media has done playing. - DONE
    # the alias is useless, as winmm just uses the filename for other commands. I just think it looks cleaner.

    logger.debug('Starting')

    winCommand(u'open {} type waveaudio alias media'.format(sound))
    winCommand(u'play media')
    winCommand(u'set media time format milliseconds')

    durationms = winCommand(u'status media length')
    await asyncio.sleep(float(int(durationms) / 1000.0))

    print("Closing.")

    winCommand(u'close media')
    logger.debug('Returning')

    # This helped me quite a bit with some of this, and future stuff I may do:
    # https://docs.microsoft.com/en-us/windows/win32/multimedia/classifications-of-mci-commands

from platform import system
from sys import argv, exit

if system() == 'Windows':
    if __name__ == '__main__':
        play(argv[1])

else:
    logger.error("You're using an unsupported operating system. (Windows Only)")
    exit(1)
