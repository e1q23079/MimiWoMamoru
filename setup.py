from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import tkinter as tk
import tkinter.messagebox as messagebox


print("MimiWoMamoru_セットアップ")
print("設定したい音量に調節してください．")
print("設定が完了したら，yを入力してください．")
nyuryoku =input()
while(nyuryoku!="y"):
    print("エラーが発生しました．")
    print("設定が完了したら，yを入力してください．")
    nyuryoku=input()

devices = AudioUtilities.GetSpeakers()

interface =devices.Activate(IAudioEndpointVolume._iid_,CLSCTX_ALL,None)

volume = interface.QueryInterface(IAudioEndpointVolume)
print(">>Setting...")
#print("%s",volume.GetMasterVolumeLevel())

onryo = volume.GetMasterVolumeLevel()

write_file = open('setting.mom','w')
write_file.write(str(onryo))
write_file.close()


print("音量の初期値の設定が完了しました．")
print("Bluetoothデバイスの接続を行ってください")
print("接続が完了したら，yを入力してください．")
nyuryoku = input()
while(nyuryoku!="y"):
    print("エラーが発生しました．")
    print("設定が完了したら，yを入力してください．")
    nyuryoku=input()


devices = AudioUtilities.GetSpeakers()

interface =devices.Activate(IAudioEndpointVolume._iid_,CLSCTX_ALL,None)

volume = interface.QueryInterface(IAudioEndpointVolume)
print(">>Setting...")
#print("%s",volume.GetMasterVolumeLevel())

ble_onryo = volume.GetMasterVolumeLevel()


write_file_ble = open('setting_ble.mom','w')
write_file_ble.write(str(ble_onryo))
write_file_ble.close()


write_file_sa = open('setting_sa.mom','w')
write_file_sa.write("15")
write_file_sa.close()

write_file_sl = open('setting_sl.mom','w')
write_file_sl.write("0")
write_file_sl.close()

write_file_mode = open('setting_mode.mom','w')
write_file_mode.write("1")
write_file_mode.close()


print("セットアップが完了しました．")
print("終了するにはenterキーを押してください．")
text = input()
