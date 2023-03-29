import random
import tkinter as tkr
import pygame
import os

player=tkr.Tk()

player.title("Audio player")
player.geometry("250x350")

music_dir="C:\\Music"
song=os.listdir(music_dir)
leng=len(song)
i=random.randint(0,(leng-1))


playlist=tkr.Listbox(player,highlightcolor="blue")

for item in song:
    pos=0
    
    playlist.insert(pos,item)
    pos=pos+1

pygame.init()
pygame.mixer.init()

def Play():
    pygame.mixer.music.load(os.path.join(music_dir,playlist.get(tkr.ACTIVE)))
    var.set(playlist.get(tkr.ACTIVE))
    pygame.mixer.music.play()

def ExitPlayer():
    pygame.mixer.music.stop()

def Pause():
    pygame.mixer.music.pause()

def Unpause():
    pygame.mixer.music.unpause()

def Next():
    global index
    index += 1
    pygame.mixer.music.load(os.path.join(music_dir,playlist[index]))
    pygame.mixer.music.play()

Button1=tkr.Button(player,width=4,height=3, text="Play",command=Play)
Button1.pack(fill="x")
 
Button2=tkr.Button(player,width=4,height=3, text="Stop",command=ExitPlayer)
Button2.pack(fill="x")
   
Button3=tkr.Button(player,width=4,height=3, text="Pause",command=Pause)
Button3.pack(fill="x")
   
Button4=tkr.Button(player,width=4,height=3, text="Unpause",command=Unpause)
Button4.pack(fill="x")

Button5=tkr.Button(player,width=4,height=3, text="Next",command=Next)
Button5.pack(fill="x")

label1=tkr.LabelFrame(player,text="Song name")
label1.pack(fill="both",expand="yes")
var=tkr.StringVar()
songtitle=tkr.Label(player,textvariable=var)
songtitle.pack(fill="both",expand="yes")
playlist.pack(fill="both",expand="yes")

player.mainloop()