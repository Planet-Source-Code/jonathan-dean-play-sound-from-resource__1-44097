<div align="center">

## Play sound from resource


</div>

### Description

If you want to play sounds in your program, but don't want to ship extra .wav files and make sure there in the right path, or you don't want people ripping off your stored .wav files, you can play sounds from a resource file built in to your .exe file. Although Windows resources have a "WAVE" type, Visual Basic doesn't support it. The answer: use a binary resource and play the sound from memory!
 
### More Info
 
Needs a valid .wav file stored in a resource as binary. Needs the a byte array containing the binary resource.

This function uses PlaySound rather than sndPlaySound, which is what MSDN suggests.

The key here is a semi undocumented function called VarPtr, which returns a Long value that is the memory address of a variable. This has been around since the days of DOS BASIC, but it's not as useful anymore in the days of protected memory.

The beauty of playing a sound this way, besides protecting your .wav file, is you load it to memory once, and then it's there, waiting to be instantly accessed. If you play a sound from disk, Windows needs to open the file, load it, close it, and play it EVERY time you play the sound. This way, you load it once and use it as many times as you want.

No returns.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jonathan Dean](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jonathan-dean.md)
**Level**          |Beginner
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Sound/MP3](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/sound-mp3__1-45.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jonathan-dean-play-sound-from-resource__1-44097/archive/master.zip)

### API Declarations

```
'PlaySound declare
'Note that although the first value is "lpsz,"
'which should be a string, I've changed it to a
'Long to accomidate for the integer address of the
'sound.
Public Declare Function PlaySoundMem Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As Long, ByVal hModule As Long, ByVal dwFlags As Long) As Long
'SND_ASYNC: return to program immediatly
'SND_NOWAIT: don't wait for the sound driver to
'become available if it's busy, return immediatly
'SND_NODEFAULT: don't play the default sound if
'your sound is unable to be played
'SND_MEMORY: play the sound from memory
'SND_NOSTOP: don't stop a currently playing sound
Public Const SND_ASYNC = &H1
Public Const SND_MEMORY = &H4
Public Const SND_NODEFAULT = &H2
Public Const SND_NOSTOP = &H10
Public Const SND_NOWAIT = &H2000
```


### Source Code

```
'Declare a byte array to put the sound in
Dim Sound() as Byte
'Load the binary resource into the byte array
'101 is the resource identifier of your sound
'"CUSTOM" is the resource type to use (CUSTOM is
'the default for binary)
'
'LoadResData automatically redims the variable
'so it's the right size
Sound = LoadResData(101, "CUSTOM")
'Play the sound
Call PlaySoundMem(VarPtr(Sound(0)), 0, SND_NOWAIT Or SND_NODEFAULT Or SND_MEMORY Or SND_ASYNC Or SND_NOSTOP)
'Clean up memory
'You wouldn't do this right away if you want to
'play the sound over and over
Redim Sound(0)
```

