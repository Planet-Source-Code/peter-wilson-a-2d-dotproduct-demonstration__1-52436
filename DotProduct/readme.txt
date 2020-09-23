A 2D DotProduct Demonstration
Copyright © 2004 Peter Wilson - All Rights Reserved.
====================================================
This application is a simple and clear demonstration of the DotProduct in two dimensions. The DotProduct is almost always used by game programmers, especially in shooting games like Unreal Tournament. The DotProduct allows the programmers to decide is the user is pointing his weapon at a monster, a pickup or a door. The DotProduct is also used for Back-Face culling and lighting effects. You can even use the DotProduct for real world physics calculations like "Newton's Conservation of Momentum". In fact, this is the real reason I created this little demo. I'm currently working on my Asteroids game (also on PSC) and wanted the Asteroids to bounce off each other convincingly. Since angles were involved, I figured I could probably use the DotProduct to save on calculations. I needed a little refresher course in the DotProduct, thus this project was born. If you are not interested in the DotProduct, you may be interested in the Splash screen. It has a built in Ant Simulator which probably deserves it's own submission. 

If you need answers to any of the following questions, then you need to understand the "DotProduct"

	*  Which faces are pointing at the camera? (ie. Hidden line removal for 3D applications)
	*  Which faces are pointing at the Sun/Light Source? 
	*  Is the monster looking at the player? 
	*  Is the player facing the monster? 
	*  Is the player's gun pointing at the monster, if so, how accurate is the player's aim? 
	*  Which object is the player looking at? 
	*  What angle is the light hitting the surface of my object? 
	*  Which two objects are on a collision course? 

As you can see the Dot Product has many application uses. It is useful for determining the difference in angle between two vectors; however the dot-product does NOT return angles, but rather the cosine value of the angle. If you think about it, then this is probably all you'll need. There is often a temptation to know the exact angle of something; Like...

	Q. What angle is the light hitting a surface on my helicopter?"

	A. Ok, so let's say I give you an answer of 45°, now what are you going to do with it? You'll probably want to illuminate the surface depending on this angle of 45°. Assuming our light source is bright white RGB(255, 255, 255) then we'll need to adjust these values accordingly like so, RGB(255*Cosine(45°), 255*Cosine(45°), 255*Cosine(45°))  But wait a second?!?! Why bother knowing the Angle of 45° if we're just going to convert it to a cosine value? Why not just start with a cosine value and keep it that way the whole time?  This is what the dot-product is about. It doesn't give you the angle - it bypasses this step. "It gives you the cosine of the angle." See? We never needed to know the angle at all! Yeah... yeah... you probably STILL want to know the angle right?!? Just remember, knowing the angle is probably the slow way of doing things, and thus isn't used at all by professionals.


Mouse Controls
==============
Move the vectors around the screen using the mouse.
	Left-Button	:	This controls vector 1.
	Right-Button	:	This controls vector 2.


Notes about the "Ant Simulator"
===============================
When an active ant touches a non-active ant, they both start moving again (just like real ant colonies). Adjust the number of ants in the "Reset" routine to increase or decrease their numbers. Click anywhere in the form to Reset. I got the idea from watching a David Suzuki documentary about how large populations of Ants automatically organize themselves in a pattern, but smaller any colonies do not. 


Version History
===============
March 2004	: * Initial release.


Feedback
========
If you've got any questions, praise or comments then send me an e-mail. If you feel so inclined to vote for this code on Planet-Source-Code, then that would be good too although not necessary.


Planet Source Code
==================
Here is a list of my other submissions. Just search for "Peter Wilson".
	*  A 2D Asteroids Game
	*  A 2D game - Froggies, a game of leap frog.
	*  A 2D Rotation Demo using SIN() and COS()
	*  A 2D Rotation Demo v2.0
	*  A 2D Rotation Lesson - Fly a UFO
	*  A 3D Lesson v2, Very Simple
	*  A 3D Lesson v3.1, Moderate
	*  A 3D Lesson v4, Advanced
	*  A 3D Studio v6.0 beta
	*  A collision avoidance system for games using DotProduct.
	*  A Matrix Multiplication Lesson using the game Asteroids
	*  A Simple Solar System Simulator, v1.0
	*  A Vacuum Fluorescent Display Simulator v1.0
	*  Convert Fonts to Vector Graphics using GetGlyphOutline.
	*  RGB Colour Wheel
	*  TechniColor Mouse Trails
	*  TechniColor Mouse Trails v 2


Cheers,

Peter Wilson
peter@midar.com
http://dev.midar.com/

