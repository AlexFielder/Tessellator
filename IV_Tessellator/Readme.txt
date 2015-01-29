	How to Register/Unregister 
	=======================

	1) Build Project;

	2) Copy add-in dll file to one of following locations: 
		a) Anywhere, then *.addin file <Assembly> setting should be updated to the full path including the dll name
		b) Inventor <InstallPath>\bin\ folder, then *.addin file <Assembly> setting should be the dll name only: <AddInName>.dll
		c) Inventor <InstallPath>\bin\XX folder, then *.addin file <Assembly> setting shoule be a relative path: XX\<AddInName>.dll

	3) Copy.addin manifest file to one of following locations:
		a) Inventor Version Dependent
		Windows XP:
			C:\Documents and Settings\All Users\Application Data\Autodesk\Inventor 2012\Addins\
		Windows7/Vista:
			C:\ProgramData\Autodesk\Inventor 2012\Addins\

		b) Inventor Version Independent
		Windows XP:
			C:\Documents and Settings\All Users\Application Data\Autodesk\Inventor Addins\
		Windows7/Vista:
			C:\ProgramData\Autodesk\Inventor Addins\

		c) Per User Override
		Windows XP:
			C:\Documents and Settings\<user>\Application Data\Autodesk\Inventor 2012\Addins\
		Windows7/Vista:
			C:\Users\<user>\AppData\Roaming\Autodesk\Inventor 2012\Addins\

	4) Startup Inventor, the AddIn should be loaded

	To unregister the AddIn, remove the Autodesk.<AddInName>.Inventor.addin from above mentioned .addin manifest file locations directly.

	The goal of this AddIn is to allow the user to create an Inventor Part whose internal structure is not hollow, or a shell but a structure of varying density akin to
	this: http://bit.ly/wxuNXr

	So far we can do the following:

	1) pick a profile.
	
	Still to implement:
	Measure the height of the object
	Use an existing or create a new Workplane
	create a bunch of offset workplanes based on above
	create a sketch on each workplane to host sketch points
	place a series of iFeature spheres at each of the sketch points from above
	subtract iFeatures from part file

	I will update this file as features are created.
	