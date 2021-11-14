import asyncio
import logging
from typing import Any

logger = logging.getLogger(__name__)


class PlaysoundException(Exception):
    pass

def winCommand(*command):
    # Because of multiple arguments (*), they needed to be joined
    command = ' '.join(command)

    '''
        Utilizes windll.winmm. Tested and known to work with MP3 and WAVE on
        Windows 7 with Python 2.7. Probably works with more file formats.
        Probably works on Windows XP thru Windows 10. Probably works with all
        versions of Python.

        Inspired by (but not copied from) Michael Gundlach <gundlach@gmail.com>'s mp3play:
        https://github.com/michaelgundlach/mp3play

        I never would have tried using windll.winmm without seeing his code.
        
        - Taylor
        
        More help about MCI below, MCI can be confusing, and because
        MCI is no longer used (to the masses anyway) documentation can be hard to find. I STRONGLY recommend
        you read up on MCI & C Data types before editing this code. (If you already know, then carry on :) )
        
        - Austin Ares
        
        Documentation Below
        
        https://docs.microsoft.com/en-us/windows/win32/multimedia/classifications-of-mci-commands
        https://github.com/michaelgundlach/mp3play/blob/master/mp3play/windows.py
        https://www.geeksforgeeks.org/play-sound-in-python/
        https://www.vbforums.com/showthread.php?276466-mci-error-codes
        https://lawlessguy.wordpress.com/2016/02/10/play-mp3-files-with-python-windows/
        https://web.archive.org/web/20190913004846/https://docs.microsoft.com/en-us/previous-versions/ms712587(v=vs.85)?redirectedfrom=MSDN
        https://www.khroma.eu/digimedia/htmlhelp/html/hs98.htm
        https://pclogo.fandom.com/wiki/MCI
        https://docs.microsoft.com/en-us/previous-versions/dd757161(v=vs.85)
        
    '''

    from ctypes import create_unicode_buffer, windll, wintypes

    # The lists here contain C data types, used for returning data & parameters for mciSendStringW()
    windll.winmm.mciSendStringW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.UINT, wintypes.HANDLE]
    windll.winmm.mciGetErrorStringW.argtypes = [wintypes.DWORD, wintypes.LPWSTR, wintypes.UINT]

    bufferLength = 600
    buffer = create_unicode_buffer(bufferLength)
    errorCode = int(windll.winmm.mciSendStringW(command, buffer, bufferLength - 1, 0))  # use widestring version of the function

    if errorCode:
        errorBuffer = create_unicode_buffer(bufferLength)
        windll.winmm.mciGetErrorStringW(errorCode, errorBuffer, bufferLength - 1)  # use widestring version of the function

        raise PlaysoundException("Error Code: {}, {}".format(errorCode, errorBuffer.value))

    return buffer.value

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
    responce = winCommand(u'{}'.format(' '.join(command)))
    return responce

'''
    Seeks to the time specified.
'''
def seek(seconds: int):
    winCommand(u'seek media to {}'.format(seconds*1000)) # Convert seconds to Miliseconds, which is what MCI is configured to use in this situation.
    winCommand(u'play media')

async def play(sound):

    # All changes here are not done, I just need to winCommand(u'close media') after the media has done playing. - DONE
    # the alias is useless, as winmm just uses the filename for other commands. I just think it looks cleaner.

    logger.debug('Starting')

    '''
        open: MCI command that specifies to open a file.
        type waveaudio: We are playing WAVE Audio, usage / help on it can be found in the winCommand Function. 
        alias media: the "media" part of that can be anything, just used for readability, and ease-of-use.
    '''
    winCommand(u'open {} alias media'.format(sound))
    winCommand(u'play media')
    winCommand(u'set media time format milliseconds')

    durationms = winCommand(u'status media length')

    # Using asyncio.sleep() because time.sleep() would block this task.
    await asyncio.sleep(float(int(durationms) / 1000.0))

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
