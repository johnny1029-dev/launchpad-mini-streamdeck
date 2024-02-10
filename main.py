import mido
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import winsdk.windows.media.control as wmc
import keyboard
import asyncio
import threading
import time

dimRed = 1
brightRed = 2
dimGreen = 80
brightGreen = 120
yellow = 117
orange = 127  # custom colors

colorPalette = {"playing": brightGreen,
                "paused": orange,
                "stopped": brightRed,
                "volSlider": yellow,
                "volButtons": yellow,
                "lockedMain": brightRed,
                "lockedSlider": dimRed,
                "unlockedMain": brightGreen,
                "unlockedSlider": dimGreen,
                "volOutOfRange": brightRed,
                "skipUnavailable": dimRed,
                "skipAvailable": dimGreen,
                "muted": brightRed,
                "unMuted": dimGreen,
                "playerUnavailable": dimRed}  # custom color palette; the names are pretty self-explanatory

volControl = False
unlocked = True
volCurves = [0, 10, 20, 30, 50, 60, 85, 100]  # volumes for 'slider' in percent
colors = [2, 2, 127, 2, 120, 127]  # 2 => red 127 => yellow/orange 120 => green (lp mini mk2)
bindings = {115: "previous track",
            116: "play/pause",
            117: "next track",
            118: "volume down",
            119: "volume up",
            120: "volume mute"}


async def get_media_session():
    s = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
    return s


def refresh_session(_a, _b):
    global session
    session = sessions.get_current_session()


def handle_input():
    global volControl, unlocked

    for message in inputMIDI:
        if unlocked:
            if message.type == "note_on":
                if message.velocity != 127:
                    continue
                for i in range(115, 121):
                    if message.note in [112, 96, 80, 64, 48, 32, 16, 0]:
                        if volControl:
                            interface.SetMasterVolumeLevelScalar(float(volCurves[int((112 - message.note) / 16)]) / 100.0, None)
                        continue
                    if message.note == i:
                        keyboard.press_and_release(bindings[i])
            if message.type == "control_change" and message.control == 104 and message.value == 127:
                if volControl:
                    volControl = False
                    outputMIDI.send(mido.Message('control_change', control=104, value=colorPalette["lockedSlider"]))
                    continue
                volControl = True
                outputMIDI.send(mido.Message('control_change', control=104, value=colorPalette["unlockedSlider"]))
        if message.type == "control_change" and message.control == 111 and message.value == 127:
            if unlocked:
                unlocked = False
                outputMIDI.send(mido.Message('control_change', control=111, value=colorPalette["lockedMain"]))
                continue
            unlocked = True
            outputMIDI.send(mido.Message('control_change', control=111, value=colorPalette["unlockedMain"]))


def update():
    global vol
    pi = session.get_playback_info()
    outputMIDI.send(mido.Message('note_on', note=116, velocity=colors[pi.playback_status]))
    col1, col2, col3, col4, col5 = colorPalette["volOutOfRange"], colorPalette["volOutOfRange"], colorPalette["muted"], colorPalette["skipUnavailable"], colorPalette["skipUnavailable"]

    if vol > 0:
        col1 = colorPalette["volButtons"]
    if vol < 100:
        col2 = colorPalette["volButtons"]
    if interface.GetMute() == 0:
        col3 = colorPalette["unMuted"]
    if pi.controls.is_previous_enabled:
        col4 = colorPalette["skipAvailable"]
    if pi.controls.is_next_enabled:
        col5 = colorPalette["skipAvailable"]

    outputMIDI.send(mido.Message('note_on', note=118, velocity=col1))
    outputMIDI.send(mido.Message('note_on', note=119, velocity=col2))
    outputMIDI.send(mido.Message('note_on', note=120, velocity=col3))
    outputMIDI.send(mido.Message('note_on', note=115, velocity=col4))
    outputMIDI.send(mido.Message('note_on', note=117, velocity=col5))


def volume_slider():
    global vol
    for i in range(0, 8):
        if vol >= volCurves[i]:
            outputMIDI.send(mido.Message('note_on', note=112 - (16 * i), velocity=colorPalette["volSlider"]))
        else:
            outputMIDI.send(mido.Message('note_off', note=112 - (16 * i), velocity=127))


if __name__ != "__main__":
    exit()

outputMIDI = mido.open_output("Launchpad Mini 1")  # change if you want to use other devices
inputMIDI = mido.open_input("Launchpad Mini 0")

outputMIDI.send(mido.Message('control_change', control=104, value=colorPalette["lockedSlider"]))
outputMIDI.send(mido.Message('control_change', control=111, value=colorPalette["unlockedMain"]))

sessions = asyncio.run(get_media_session())
session = None
sessions.add_current_session_changed(refresh_session)
refresh_session(None, None)
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None).QueryInterface(IAudioEndpointVolume)

inputThread = threading.Thread(target=handle_input)
inputThread.start()

try:
    while True:
        vol = round(interface.GetMasterVolumeLevelScalar() * 100)
        volume_slider()
        if session is not None:
            update()
        else:
            outputMIDI.send(mido.Message('note_on', note=115, velocity=colorPalette["playerUnavailable"]))
            outputMIDI.send(mido.Message('note_on', note=116, velocity=colorPalette["playerUnavailable"]))
            outputMIDI.send(mido.Message('note_on', note=117, velocity=colorPalette["playerUnavailable"]))
            col1, col2, col3 = colorPalette["volOutOfRange"], colorPalette["volOutOfRange"], colorPalette["muted"]

            if vol > 0:
                col1 = colorPalette["volButtons"]
            if vol < 100:
                col2 = colorPalette["volButtons"]
            if interface.GetMute() == 0:
                col3 = colorPalette["unMuted"]

            outputMIDI.send(mido.Message('note_on', note=118, velocity=col1))
            outputMIDI.send(mido.Message('note_on', note=119, velocity=col2))
            outputMIDI.send(mido.Message('note_on', note=120, velocity=col3))

        time.sleep(0.2)
except KeyboardInterrupt:
    for i in range(0, 121):
        outputMIDI.send(mido.Message('note_off', note=i, velocity=127))
    outputMIDI.send(mido.Message('control_change', control=104, value=0))
    outputMIDI.send(mido.Message('control_change', control=111, value=0))
    inputThread.join(0)
    inputMIDI.close()
    outputMIDI.close()
    interface.Release()
