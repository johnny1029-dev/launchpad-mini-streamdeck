import mido
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import keyboard
import asyncio
import winsdk.windows.media.control as wmc
import threading

colors = [2, 2, 127, 2, 120, 127]   # 2 => red 127 => yellow/orange 120 => green (lp mini mk2)


async def get_media_session():
    try:
        sessions = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
        session = sessions.get_current_session()
        return session
    except OSError as e:
        if e.winerror == -2147023170:
            return 0
        else:
            print("gah damn")
            return 0


def media_state():
    session = asyncio.run(get_media_session())
    if session is None:
        return False
    try:
        return session.get_playback_info().playback_status
    except AttributeError:
        return 0


def previous_available():
    session = asyncio.run(get_media_session())
    if session is None:
        return False
    try:
        return session.get_playback_info().controls.is_previous_enabled
    except AttributeError:
        return False


def next_available():
    session = asyncio.run(get_media_session())
    if session is None:
        return False
    try:
        return session.get_playback_info().controls.is_next_enabled
    except AttributeError:
        return False


def volume():
    global interface
    volume = interface.QueryInterface(IAudioEndpointVolume)
    return round(volume.GetMasterVolumeLevelScalar() * 100)


def get_mute():
    global interface
    return interface.QueryInterface(IAudioEndpointVolume).GetMute()


def check_key(message, key):
    return message.note == key and message.velocity == 127


def handle_input():
    for message in inputMIDI:
        if check_key(message, 116):
            keyboard.press_and_release('play/pause')
        if check_key(message, 118):
            keyboard.press_and_release('volume down')
        if check_key(message, 119):
            keyboard.press_and_release('volume up')
        if check_key(message, 120):
            keyboard.press_and_release('volume mute')
        if check_key(message, 115):
            keyboard.press_and_release('previous track')
        if check_key(message, 117):
            keyboard.press_and_release('next track')


def update():
    outputMIDI.send(mido.Message('note_on', note=116, velocity=colors[media_state()]))
    vol = volume()
    col1, col2, col3, col4, col5 = 2, 2, 2, 2, 2
    if vol > 0:
        col1 = 127
    if vol < 100:
        col2 = 127
    if get_mute() == 0:
        col3 = 120
    if previous_available():
        col4 = 120
    if next_available():
        col5 = 120
    outputMIDI.send(mido.Message('note_on', note=118, velocity=col1))
    outputMIDI.send(mido.Message('note_on', note=119, velocity=col2))
    outputMIDI.send(mido.Message('note_on', note=120, velocity=col3))
    outputMIDI.send(mido.Message('note_on', note=115, velocity=col4))
    outputMIDI.send(mido.Message('note_on', note=117, velocity=col5))


devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
outputMIDI = mido.open_output("Launchpad Mini 1") # change if you want to use other devices
inputMIDI = mido.open_input("Launchpad Mini 0")
inputThread = threading.Thread(target=handle_input)
inputThread.start()
print(get_mute())
try:
    while True:
        update()
except KeyboardInterrupt:
    outputMIDI.send(mido.Message('note_off', note=116, velocity=127))
    outputMIDI.send(mido.Message('note_off', note=118, velocity=127))
    outputMIDI.send(mido.Message('note_off', note=119, velocity=127))
    outputMIDI.send(mido.Message('note_off', note=120, velocity=127))
    outputMIDI.send(mido.Message('note_off', note=115, velocity=127))
    outputMIDI.send(mido.Message('note_off', note=117, velocity=127))
    inputThread.join(0)
    inputMIDI.close()
    interface.Release()
