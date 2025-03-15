from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import time
import tkinter as tk
import tkinter.messagebox as messagebox
import math

ok = 4

#設定データの読み取り

try:
    with open('setting.mom','r') as read_file:
        data = read_file.read()
        ok=1
except FileNotFoundError:
        ok=0


try:
    with open('setting_ble.mom','r') as read_file_ble:
        data_ble = read_file_ble.read()
        ok=1
except FileNotFoundError:
        ok=0

try:
    with open('setting_sa.mom','r') as read_file_sa:
        data_sa = read_file_sa.read()
        ok=1
except FileNotFoundError:
        ok=0

try:
    with open('setting_sl.mom','r') as read_file_sl:
        data_sl = read_file_sl.read()
        ok=1
except FileNotFoundError:
        ok=0

try:
    with open('setting_mode.mom','r') as read_file_mode:
        data_mode = read_file_mode.read()
        ok=1
except FileNotFoundError:
        ok=0
        
    

if(ok==0):
    #print("セットアップが完了していません．")
    tk.Tk().withdraw()
    messagebox.showinfo('MimiWoMamoru','セットアップが完了していません．')


if(ok==1):    



    read_file.close()
    read_file_ble.close()
    read_file_sa.close()
    read_file_sl.close()
    read_file_mode.close()

    devices = AudioUtilities.GetSpeakers()

    interface =devices.Activate(IAudioEndpointVolume._iid_,CLSCTX_ALL,None)

    volume = interface.QueryInterface(IAudioEndpointVolume)
    #print("%s",volume.GetMasterVolumeLevel())

    zenonryo = volume.GetMasterVolumeLevel()

    flag=0

    while(flag==0):

        devices = AudioUtilities.GetSpeakers()

        interface =devices.Activate(IAudioEndpointVolume._iid_,CLSCTX_ALL,None)

        volume = interface.QueryInterface(IAudioEndpointVolume)


        time.sleep(1)
        nowonryo = volume.GetMasterVolumeLevel()
        #print("音量")
        #print(int(nowonryo))
        #print("音量差")
        #print(int(nowonryo)-int(zenonryo))
        integer = math.floor(float(data_ble))
        #print(integer)
        #print("datasa:"+data_sa)
        if((int(nowonryo)-int(zenonryo)>int(data_sa)) and (int(nowonryo)>(integer-3))):
            
            #data_saは15が標準
            #print("大規模変化検知")
            volume.SetMasterVolumeLevel(float(data),None)
            if(int(data_mode)==2):
                time.sleep(int(data_sl))
                volume.SetMasterVolumeLevel(float(data),None)
            #[-23が標準]
        zenonryo = nowonryo
        
