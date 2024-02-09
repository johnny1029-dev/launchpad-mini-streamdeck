import mido
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import keyboard
import asyncio
import winsdk.windows.media.control as wmc
import threading
import time

volCurves = [0, 15, 30, 45, 60, 75, 90, 100]  # volumes for 'slider' in percent
colors = [2, 2, 127, 2, 120, 127]             # 2 => red 127 => yellow/orange 120 => green (lp mini mk2)
bindings = {115: "previous track",
            116: "play/pause",
            117: "next track",
            118: "volume down",
            119: "volume up",
            120: "volume mute"}


async def get_media_session():
    try:
        sessions = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
        return sessions
    except OSError as e:
        if e.winerror != -2147023170:
            print(e)
        return None


def media_state():
    global session
    if session is None:
        return False
    return session.get_playback_info().playback_status


def previous_available():
    global session
    if session is None:
        return False
    return session.get_playback_info().controls.is_previous_enabled


def next_available():
    global session
    if session is None:
        return False
    return session.get_playback_info().controls.is_next_enabled


def volume():
    global interface
    return round(interface.QueryInterface(IAudioEndpointVolume).GetMasterVolumeLevelScalar() * 100)


def get_mute():
    global interface
    return interface.QueryInterface(IAudioEndpointVolume).GetMute()


def check_key(message, key):
    return message.note == key and message.velocity == 127


def handle_input():
    for message in inputMIDI:
        for i in range(115, 121):
            if check_key(message, i):
                keyboard.press_and_release(bindings[i])


def update():
    outputMIDI.send(mido.Message('note_on', note=116, velocity=colors[media_state()]))
    vol = volume()
    col1, col2, col3, col4, col5 = 2, 2, 2, 1, 1
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


def refresh_session(_a, _b):
    global session
    session = sessions.get_current_session()


sessions = asyncio.run(get_media_session())
session = None
sessions.add_current_session_changed(refresh_session)
refresh_session(None, None)
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
outputMIDI = mido.open_output("Launchpad Mini 1")  # change if you want to use other devices
inputMIDI = mido.open_input("Launchpad Mini 0")
inputThread = threading.Thread(target=handle_input)
inputThread.start()
try:
    while True:
        if session is not None:
            update()
        time.sleep(1)
except KeyboardInterrupt:
    for i in range(115, 121):
        outputMIDI.send(mido.Message('note_off', note=i, velocity=127))
    inputThread.join(0)
    inputMIDI.close()
    interface.Release()
