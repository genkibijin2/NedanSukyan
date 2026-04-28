'//=======================INFO PAGE========================\\
Private animationTimer as New Timer
'New timer for delaying an animation that will play, so that it can play after the form loads automatically.
'Private intervalCount as Integer
'Counts each step of the timer

Private Sub Form_Load()
	animationTimer.Interval = 200
	animationTimer.Enabled = true
	infoBox.Text = "A collection of companies that supply parts and materials to us at Euroglaze. This program is designed to be a useful way for me to keep track of the different methods required to price check any invoices, as it's portable and on my desk.\nIt offers the suppliers locational data, and then the various individual steps I need to take to check their invoices.\nIt also functions as a way for me to master VB and Handheld basic.\nMy name's Kendall Price and I made this in 2025...\nhttps://\nkendalls.garden\n\n\n\n"
End Sub

Private Sub returnButton_Click()
	dim f as new launchScreen
  f.Show hbFormModeless+hbFormGoto
End Sub

Private Sub Form_Unload()
	animationTimer.Enabled = false
End Sub

Private Sub animationTimer_Timer()
	animationTimer.Interval = 100000
	playAnimation
End Sub

Private Sub playAnimation()
		Dim frameInterval As Integer
		frameInterval = 30
		Dim gifLoops As Integer
			'//===EXCESSIVELY BULKY AND UNOPTIMIZED ANIMATION=====\\
			For gifLoops = 1 To 3
					Set animationBox.Image = animationFrame2
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame3
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame4
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame5
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame6
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame7
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame8
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame9
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame10
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame11
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame12
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame13
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame14
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame15
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame16
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame17
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame18
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame19
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame20
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame21
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame22
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame23
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame24
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame25
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame26
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame27
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame28
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame29
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame30
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame31
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame32
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame33
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame34
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame35
					Sleep(frameInterval)
					Set animationBox.Image = animationFrame36
					Sleep(frameInterval)
			Next gifLoops
			animationTimer.Enabled = false
			spinner.Visible = true
End Sub
'\\========================================================//

Private Sub spinner_Click()
	playAnimation
End Sub
