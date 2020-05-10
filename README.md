# TLSA Engine (Visual Basic 6)
#### My last version of TLSA Engine developed in Visual Basic 6.0 with dx_lib32 2.1 at mid 2010, a custom 2D cinematic platform engine. Include SDK tools and some test projects.

![tlsaengine_vb6_editor01.jpg](https://i0.wp.com/portfolio.visualstudioex3.com/wp-content/uploads/2010/05/tlsaengine_vb6_editor01.jpg)

##### Development date/period: 
* 2005 - 2010

##### History:
A 2D Game Engine based on __dx_lib32 Project__, with the purpose of develop a 2D platform game like __Flashback__ or __Another World__.

A little resume of their features in latest versions:

* Component Oriented Engine, trying to simulate the __XNA__ architecture.
* __2D Graphic Engine__, multilayer sprite based, with simple effect system based on __Directx 8.1__ fixed pipeline, applicable to individual sprites or to entire scene (the final scene are a transformable canvas with support for all sprite effects, position, rotation and scale transformations), implementing a sprite control point map system (similar as how __Div Game Studio__ implementing in his sprite system) to manage easily multiple textures and transformations in a nested object group (to create complex animations, based in multiple pieces, with independent sprite animations), simple camera system (with support to define multiple scene cameras, to switch between them easily using paths or animations with scales and rotations).
* __2D Audio Engine__, with support for basic realtime standard effects (non parametrizables), and spatial system to simulate distances and position listeners changing the stereo volume level of the sound effects, and a basic mutichannel mixer.
* __Basic Input System__, based on actions, which can define multiple input (keyboard, mouse and joysticks or gamepads), and complete support for joysticks and gamepads, via __DirectInput 8__ and __XInput 1.3__ (for fully support __XBox 360 Gamepads__), with basic __Force Feedback__ support (to simulate the __XInput__ rumble system in compatible joysticks using constant force effect).
* __Basic Physics Collision engine__, with multiple layer collision system, world partition areas, raycaster and force emitters (to simulate explosions or black holes forces).
* __WYSIWYG Level Editor__ with flow controls (play, pause and restart scene during the debug), scene physics designer, an audio areas designer (for applying effects and emitters) that uses the physics defined in the scene, and visual debugger.
* Some tools in the SDK like the __Input Editor__ to create profiles input files, with define actions and her input controls, to import in the game engine easily, and the __Tile Studio__, a simple but complete editor to define tile sheets and sprite sheets with irregular sizes, control point definitions, and animation sequences, with animation previsualizer.

The first versions of the engine development are from 2005 and 2006. The last version, reprogrammed from scratch, started development during the summer of 2009, and during until the last built, in summer of 2010.

This game engine is not finished, because the complexity to develop a project as this in __Visual Basic 6.0__. The game engine was used in few projects, mostly a prototypes and gamejams.

##### Notes:
* The comment lines in code are in spanish.

##### Related links:
* dx_lib32 Project: http://portfolio.visualstudioex3.com/2006/02/25/dxlib32-project/
* TLSA Engine: http://portfolio.visualstudioex3.com/2010/07/30/tlsa-engine-vb6/
* Marius Watz Java implementation for 2D line intersection function: http://workshop.evolutionzone.com/2007/09/10/code-2d-line-intersection/

![tlsaengine_01.jpg](https://i0.wp.com/portfolio.visualstudioex3.com/wp-content/uploads/2010/05/tlsaengine_01.jpg?w=414&ssl=1) ![tlsaengine_02.jpg](https://i0.wp.com/portfolio.visualstudioex3.com/wp-content/uploads/2010/05/tlsaengine_02.jpg?w=414&ssl=1)
![tlsaengine_03.jpg](https://i2.wp.com/portfolio.visualstudioex3.com/wp-content/uploads/2010/05/tlsaengine_03.jpg?w=414&ssl=1) ![tlsaengine__vb6_colidetest_00.png](https://i2.wp.com/portfolio.visualstudioex3.com/wp-content/uploads/2010/05/tlsaengine__vb6_colidetest_00.png?w=414&ssl=1)
![tlsaengine_vb6_editor00.jpg](https://i2.wp.com/portfolio.visualstudioex3.com/wp-content/uploads/2010/05/tlsaengine_vb6_editor00.jpg?w=414&ssl=1) ![tlsaengine_vb6_editor03.jpg](https://i1.wp.com/portfolio.visualstudioex3.com/wp-content/uploads/2010/05/tlsaengine_vb6_editor03.jpg?w=414&ssl=1)
![tlsaengine_vb6_tileeditor_00.png](https://i0.wp.com/portfolio.visualstudioex3.com/wp-content/uploads/2010/05/tlsaengine_vb6_tileeditor_00.png?w=832&ssl=1) 
