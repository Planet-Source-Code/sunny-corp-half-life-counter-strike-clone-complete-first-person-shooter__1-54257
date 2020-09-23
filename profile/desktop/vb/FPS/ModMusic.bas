Attribute VB_Name = "ModSound"
Option Explicit

'set the main direct sound object
Dim DX_DSound As DirectSound

'for each sound effect there is a different DS buffer
Global Walk As DirectSoundBuffer
Global MShoot As DirectSoundBuffer
Global BShoot As DirectSoundBuffer
Global Object As DirectSoundBuffer
Global Objective1 As DirectSoundBuffer
Global EndA As DirectSoundBuffer
Global Yoda1 As DirectSoundBuffer
Global Yoda2 As DirectSoundBuffer
Global Darth_Vader1 As DirectSoundBuffer
Global Maul As DirectSoundBuffer
Global Qui As DirectSoundBuffer
Global Soldier1 As DirectSoundBuffer
Global Soldier3 As DirectSoundBuffer
Global Soldier4 As DirectSoundBuffer
Global Player As DirectSoundBuffer
Global Droid As DirectSoundBuffer
Global Saber_on As DirectSoundBuffer
Global Saber_BHit As DirectSoundBuffer
Global Saber_Hit As DirectSoundBuffer
Global Saber_Swing1 As DirectSoundBuffer
Global saber_Swing2 As DirectSoundBuffer

'for each sound there is corresponding text
Global objective As String
Global Cyborg1_Text As String
Global Cyborg2_Text As String
Global Yoda1_Text As String
Global Yoda1_Texta As String
Global Yoda1_Textb As String
Global Yoda2_Text As String
Global Yoda2_Texta As String
Global Yoda2_Textb As String
Global Yoda2_Textc As String
Global Yoda3_Text As String
Global Dart1_Text As String
Global Qui_Text As String
Global Maul_Text As String
Global Computer1_Text As String
Global Computer1_Texta As String
Global Soldier1_Text As String
Global Computer2_Text As String
Global Computer2_Texta As String
Global Soldier2_Text As String
Global Soldier2_Texta As String
Global Soldier3_text As String
Global Droid_Text As String
Global Player_Text As String
Global Yoda3_Texta As String

Global Cyborg_Sp As Boolean 'used to check if cyborg is still speaking

'directsound options
Dim DsDesc As DSBUFFERDESC
Dim DsWave As WAVEFORMATEX


Public Sub DirectSound_Init()
    On Local Error Resume Next
    
    'set DS object to the actual DX7 object
    Set DX_DSound = DX_Main.DirectSoundCreate("")
    If Err.Number <> 0 Then
        MsgBox "Direct Sound Initialization Failed."
        'Call DX_Exit
    End If
    Text_Init
    'set cooperative level
    DX_DSound.SetCooperativeLevel frmMain.hWnd, DSSCL_NORMAL
    
    'set sound quality, and other propr.
    DsDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC Or DSBCAPS_CTRLPOSITIONNOTIFY
    DsWave.nFormatTag = WAVE_FORMAT_PCM  'PCM version only
    DsWave.nChannels = 2 'we need stereo sound! not mono
    DsWave.lSamplesPerSec = 22050
    DsWave.nBitsPerSample = 16 '16 bits mean higher quality but slower loading time and more memory consuming
    DsWave.nBlockAlign = DsWave.nBitsPerSample / 8 * DsWave.nChannels
    DsWave.lAvgBytesPerSec = DsWave.lSamplesPerSec * DsWave.nBlockAlign
    
    'set up each different sound effect or in some cases the conversations
    Set Walk = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/walk.wav", DsDesc, DsWave)
    Set MShoot = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Shoot.wav", DsDesc, DsWave)
    Set BShoot = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Laser.wav", DsDesc, DsWave)
    Set Object = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Conversation/solder1.wav", DsDesc, DsWave)
    Set Objective1 = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Conversation/objectives.wav", DsDesc, DsWave)
    Set EndA = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Conversation/enda.wav", DsDesc, DsWave)
    Set Yoda1 = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Conversation/Yoda_Master.wav", DsDesc, DsWave)
    Set Yoda2 = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Conversation/Yoda5.wav", DsDesc, DsWave)
    Set Player = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Conversation/Player1.wav", DsDesc, DsWave)
    Set Droid = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Conversation/Droid.wav", DsDesc, DsWave)
    Set Darth_Vader1 = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Conversation/Darthvadar2.wav", DsDesc, DsWave)
    Set Maul = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Conversation/darthmaul2.wav", DsDesc, DsWave)
    Set Qui = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Conversation/qui-gon-blah2.wav", DsDesc, DsWave)
    Set Saber_on = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Saber/saber_on.wav", DsDesc, DsWave)
    Set Saber_BHit = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Saber/ltsaberbodyhit01.wav", DsDesc, DsWave)
    Set Saber_Hit = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Saber/ltsaberhit02.wav", DsDesc, DsWave)
    Set Saber_Swing1 = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Saber/ltsaberswing06.wav", DsDesc, DsWave)
    Set saber_Swing2 = DX_DSound.CreateSoundBufferFromFile(App.Path & "/Sounds/Saber/ltsaberswingdbl01.wav", DsDesc, DsWave)
    
    'Debug.Print App.Path
End Sub

Public Sub Text_Init()
      'message which go with sound files
      objective = "Dispatcher: There is a hostage situation, FIND the Hostage, and get out of there."
      Yoda1_Text = "Yoda: Got Here I have. How? I know not."
      Yoda1_Texta = "Yoda: Master Yoda I am, yes,yes. Much i've heard about you. "
      Yoda1_Textb = "Yoda: Dark side approaching fast it is, defend you must against it."
      Yoda2_Text = "Yoda: Four Jedi Knights have come, Light brings with it 2 guardians, "
      Yoda2_Textb = "Yoda: while the dark side brings two of its own kind it does. "
      Yoda2_Textc = "Yoda: Get involved in their conflict, you must not. Droids must you destroy."
      Yoda3_Text = "Yoda: Won one battle you have; victory is far away still it is. "
      Yoda3_Texta = "Yoda: Go I must but may the Force be with you."
      Cyborg1_Text = "Cyborg: You may have inflitrated our base now, but you will never escape!"
      Soldier1_Text = "Dispatcher: Quick! You have been detected; KILL all opposition forces."
      Cyborg2_Text = "Cyborg: Don't get too happy, more are coming, you won't get out alive."
      Computer1_Text = "Computer: New Objective; Free the HOSTAGE, and ELIMINATE any opposition,..."
      Computer1_Texta = "Computer: Your're licensed to kill.   "
      Computer2_Text = "Computer:Objective complete, hold position for extraction point coordinates, ..."
      Computer2_Texta = "Computer:(static)... We're loosing contact; unknown error! ....(static)... "
      Soldier2_Text = "Soldier: Sir, we're loosing...(static)..all communications,... "
      Soldier2_Texta = "Soldier:...something big is happening in your vicinity.....(static)... "
      Player_Text = "Player: Who are you?"
      Soldier3_text = "Soldier: Sir we've gained back communictations, transport approaching your location."
      Dart1_Text = "Darth Vader: I've.... failed.... you.... master."
      Qui_Text = "Qui-Gon-Jin: May,... the...force...be....with...you.....I must go now....."
      Maul_Text = "Darth Maul: I've......lost.....WE WILL COME AGAIN!!!!"
End Sub

Public Sub DSSound_Exit()
    'just to make sure the sound is stopped even when the esc key is pressed
    
    'stop and finsih most of these effects
    Walk.Stop
    Walk.SetCurrentPosition 0
    
    
    MShoot.Stop
    MShoot.SetCurrentPosition 0
    
    BShoot.Stop
    BShoot.SetCurrentPosition 0
    
    Object.Stop
    Object.SetCurrentPosition 0
    
    Objective1.Stop
    Objective1.SetCurrentPosition 0
    
    EndA.Stop
    EndA.SetCurrentPosition 0
End Sub
