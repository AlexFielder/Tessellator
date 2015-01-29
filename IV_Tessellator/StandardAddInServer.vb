Imports System.Collections.Generic
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
'Imports Connectivity.InventorAddin.EdmAddin
Imports Inventor

Namespace IVTessellator
    <ProgIdAttribute("IV_Tessellator.StandardAddInServer"), _
    GuidAttribute("4103d4f8-04bd-4ee2-9eb7-ede1493daa13")> _
    Public Class StandardAddInServer

        Implements Inventor.ApplicationAddInServer
        Friend Shared Instance As StandardAddInServer

        ' Inventor application object.
        'Private m_inventorApplication As Inventor.Application

#Region "ApplicationAddInServer Members"

        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate

            ' This method is called by Inventor when it loads the AddIn.
            ' The AddInSiteObject provides access to the Inventor Application object.
            ' The FirstTime flag indicates if the AddIn is loaded for the first time.

            ' Initialize AddIn members.
            ThisApplication = addInSiteObject.Application

            ' TODO:  Add ApplicationAddInServer.Activate implementation.
            ' e.g. event initialization, command creation etc.

            'initialize event handlers
            m_userInterfaceEvents = ThisApplication.UserInterfaceManager.UserInterfaceEvents
            m_applicationEvents = ThisApplication.ApplicationEvents

            AddHandler m_userInterfaceEvents.OnResetCommandBars, AddressOf Me.UserInterfaceEvents_OnResetCommandBars
            AddHandler m_userInterfaceEvents.OnEnvironmentChange, AddressOf Me.UserInterfaceEvents_OnEnvironmentChange
            AddHandler m_userInterfaceEvents.OnResetRibbonInterface, AddressOf Me.UserInterfaceEvents_OnResetRibbonInterface

            'Retrieve the GUID for this class
            Dim addInCLSID As GuidAttribute
            addInCLSID = CType(System.Attribute.GetCustomAttribute(GetType(StandardAddInServer), GetType(GuidAttribute)), GuidAttribute)

            m_addInCLSIDString = "{" & addInCLSID.Value & "}"

            'create the command button(s)

            ' Set the large icon size based on whether it is the classic or ribbon interface.
            Dim UIManager As UserInterfaceManager = ThisApplication.UserInterfaceManager

            Dim largeIconSize As Integer
            'forgot this bit for the longest time - would have broken everything. I think?
            If ThisApplication.UserInterfaceManager.InterfaceStyle = InterfaceStyleEnum.kRibbonInterface Then
                largeIconSize = 32
            Else
                largeIconSize = 24
            End If
            Dim controlDefs As ControlDefinitions = ThisApplication.CommandManager.ControlDefinitions
            '' define 2 more icons for the screenshot tool.
            Dim HTSmallPicture As IPictureDisp = ToIPictureDisp(New Icon(My.Resources.HT, 16, 16))
            Dim HTLargePicture As IPictureDisp = ToIPictureDisp(New Icon(My.Resources.HT, largeIconSize, largeIconSize))

            m_HTButtonDef = controlDefs.AddButtonDefinition("Tessellate!", "CH2MTessellator", CommandTypesEnum.kShapeEditCmdType, _
                                                            m_addInCLSIDString, "Allow users to create Hyperbolic Tessellations within the sketch environment", _
                                                             "Allow users to create Hyperbolic Tessellations within the sketch environment", _
                                                             HTSmallPicture, HTLargePicture)
            Dim linesSmallPic As IPictureDisp = ToIPictureDisp(New Icon(My.Resources.CH2M_Globe, 16, 16))
            Dim linesLargePic As IPictureDisp = ToIPictureDisp(New Icon(My.Resources.CH2M_Globe, largeIconSize, largeIconSize))

            m_LinesButtonDef = controlDefs.AddButtonDefinition("Lines!", "CH2MLines", CommandTypesEnum.kShapeEditCmdType, _
                                                               m_addInCLSIDString, "Adds some lines!", "Click this to add some simple lines to a sketch", _
                                                               linesSmallPic, linesLargePic)

            If firstTime Then
                'Add the button to the part features toolbar
                Dim userInterfaceMgr As UserInterfaceManager
                userInterfaceMgr = ThisApplication.UserInterfaceManager

                Dim interfaceStyle As InterfaceStyleEnum = userInterfaceMgr.InterfaceStyle

                'Create the UI for classic interface
                If interfaceStyle = InterfaceStyleEnum.kClassicInterface Then
                    'Get the commandbars collection
                    Dim commandBars As CommandBars
                    commandBars = userInterfaceMgr.CommandBars

                    'Get the part sketch toolbar
                    Dim partSketchToolbar As Inventor.CommandBar
                    partSketchToolbar = commandBars("PMxPartSketchCmdBar")

                    'Add our buttons to the toolbar
                    partSketchToolbar.Controls.AddButton(m_HTButtonDef)
                    partSketchToolbar.Controls.AddButton(m_LinesButtonDef)
                    'MessageBox.Show("The Tessellator Method uses the ribbon interface only!")
                Else
                    'Get the ribbons associated with part documents
                    Dim ribbons As Ribbons = userInterfaceMgr.Ribbons
                    Dim partRibbon As Ribbon = ribbons("Part")

                    'Get the tabs associated with the Sketch ribbon
                    Dim ribbonTabs As RibbonTabs = partRibbon.RibbonTabs
                    Dim sketchTab As RibbonTab = ribbonTabs("id_TabSketch")

                    'Get the tabs associated with part ribbon
                    'Dim ribbonTabs As RibbonTabs = partRibbon.RibbonTabs
                    'Dim modelRibbonTab As RibbonTab = ribbonTabs("id_TabModel")

                    'Get the panels within Sketch tab
                    Dim ribbonPanels As RibbonPanels = sketchTab.RibbonPanels
                    Dim SketchRibbonPanel As RibbonPanel = ribbonPanels("id_PanelP_2DSketchDraw")

                    'Add controls to the Sketch panel
                    Dim sketchRibbonPanelCtrls As CommandControls = SketchRibbonPanel.CommandControls

                    'Add button(s) to the Sketch panel
                    sketchRibbonPanelCtrls.AddButton(m_HTButtonDef)
                    sketchRibbonPanelCtrls.AddButton(m_LinesButtonDef)
                End If

            End If
        End Sub

        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' This method is called by Inventor when the AddIn is unloaded.
            ' The AddIn will be unloaded either manually by the user or
            ' when the Inventor session is terminated.

            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Release objects.
            Marshal.ReleaseComObject(ThisApplication)
            ThisApplication = Nothing

            System.GC.WaitForPendingFinalizers()
            System.GC.Collect()

        End Sub

        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation

            ' This property is provided to allow the AddIn to expose an API 
            ' of its own to other programs. Typically, this  would be done by
            ' implementing the AddIn's API interface in a class and returning 
            ' that class object through this property.

            Get
                Return Nothing
            End Get

        End Property

        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand

            ' Note:this method is now obsolete, you should use the 
            ' ControlDefinition functionality for implementing commands.

        End Sub

        Public Sub UserInterfaceEvents_OnResetCommandBars(ByVal commandBars As ObjectsEnumerator, ByVal context As NameValueMap)

            Try
                Dim commandBar As CommandBar
                For Each commandBar In commandBars
                    If commandBar.InternalName = "PMxPartSketchCmdBar" Then

                        'Add our button(s) back to the part features toolbar
                        commandBar.Controls.AddButton(m_HTButtonDef)
                        commandBar.Controls.AddButton(m_LinesButtonDef)
                        'commandBar.Controls.AddButton(m_rackFaceCmd.ButtonDefinition)

                        Exit Sub
                    End If
                Next

            Catch ex As System.Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

        Public Sub UserInterfaceEvents_OnEnvironmentChange(ByVal environment As Inventor.Environment, ByVal environmentState As EnvironmentStateEnum, ByVal beforeOrAfter As EventTimingEnum, ByVal context As NameValueMap, ByRef handlingCode As HandlingCodeEnum)

            Try
                Dim envInternalName As String
                envInternalName = environment.InternalName

                If envInternalName = "PMxSketchEnvironment" Then

                    ' Enable the "Rack Face" button when the part environment is activated or resumed
                    If environmentState = EnvironmentStateEnum.kActivateEnvironmentState Or environmentState = EnvironmentStateEnum.kResumeEnvironmentState Then
                        m_HTButtonDef.Enabled = True
                        m_LinesButtonDef.Enabled = True
                        'm_rackFaceCmd.ButtonDefinition.Enabled = True
                    End If

                    ' Disable the "Rack Face" button when the part environment is terminated or suspended
                    If environmentState = EnvironmentStateEnum.kTerminateEnvironmentState Or environmentState = EnvironmentStateEnum.kSuspendEnvironmentState Then
                        m_HTButtonDef.Enabled = False
                        m_LinesButtonDef.Enabled = False
                        'm_rackFaceCmd.ButtonDefinition.Enabled = False
                    End If
                End If

                handlingCode = HandlingCodeEnum.kEventNotHandled

            Catch ex As System.Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub

        Public Sub UserInterfaceEvents_OnResetRibbonInterface(ByVal context As NameValueMap)

            Try
                Dim userInterfaceMgr As UserInterfaceManager = ThisApplication.UserInterfaceManager

                'Get the ribbons associated with part documents
                Dim ribbons As Ribbons = userInterfaceMgr.Ribbons
                Dim partRibbon As Ribbon = ribbons("Part")

                'Get the tabs associated with the Sketch ribbon
                Dim ribbonTabs As RibbonTabs = partRibbon.RibbonTabs
                Dim sketchTab As RibbonTab = ribbonTabs("id_TabSketch")

                'Get the panels within Sketch tab
                Dim ribbonPanels As RibbonPanels = sketchTab.RibbonPanels
                Dim SketchRibbonPanel As RibbonPanel = ribbonPanels("id_PanelP_2DSketchDraw")

                'Add controls to the Sketch panel
                Dim sketchRibbonPanelCtrls As CommandControls = SketchRibbonPanel.CommandControls

                'Add button to the Sketch panel
                sketchRibbonPanelCtrls.AddButton(m_HTButtonDef)
                sketchRibbonPanelCtrls.AddButton(m_LinesButtonDef)
            Catch ex As System.Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub
#End Region
#Region "Event Definitions"
        ''' <summary>
        ''' Allows us to capture/work with ApplicationEvents
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared WithEvents m_applicationEvents As ApplicationEvents
        'Public Sub New()
        '    MyBase.New()
        '    m_applicationEvents = ThisApplication.ApplicationEvents
        'End Sub
        'Protected Overrides Sub Finalize()
        '    m_applicationEvents = Nothing
        '    MyBase.Finalize()
        'End Sub

        ''' <summary>
        ''' Allows us to capture/work with UserInterfaceEvents
        ''' </summary>
        ''' <remarks></remarks>
        Private WithEvents m_userInterfaceEvents As UserInterfaceEvents

        ''' <summary>
        ''' Allows us to work with/capture UserInputEvents
        ''' </summary>
        ''' <remarks></remarks>
        Private WithEvents m_InputEvents As UserInputEvents

        ''' <summary>
        ''' Allows us to capture any FileDialogEvents
        ''' </summary>
        ''' <remarks></remarks>
        Private WithEvents m_dialogueEvents As FileDialogEvents
#End Region
#Region "File-specific definitions"
        ''' <summary>
        ''' A DrawingDocument Object
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared oIDWDoc As DrawingDocument

        ''' <summary>
        ''' A Part document object.
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared oPartDoc As PartDocument

        ''' <summary>
        ''' An Assembly document object
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared oAssyDoc As AssemblyDocument

#End Region
#Region "Variable Definitions"
        ''' <summary>
        ''' Allows us to use the ThisApplication.Something method.
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared ThisApplication As Inventor.Application

        ''' <summary>
        ''' the id of the current Client?
        ''' </summary>
        ''' <remarks></remarks>
        Private m_addInCLSIDString As String

        Private m_activeSketch As PlanarSketch
        Private m_skProfile As Profile
#End Region
#Region "Button Definitions"
        ' ''' <summary>
        ' ''' Screenshot button
        ' ''' </summary>
        ' ''' <remarks></remarks>
        Private WithEvents m_HTButtonDef As ButtonDefinition
        ''' <summary>
        ''' Creates some abitrary lines - shouldn't run before we can crawl!
        ''' </summary>
        ''' <remarks></remarks>
        Private WithEvents m_LinesButtonDef As ButtonDefinition


#End Region
#Region "Methods"

        Public Sub m_HTButtonDef_OnExecute(ByVal Context As NameValueMap) Handles m_HTButtonDef.OnExecute
            'MessageBox.Show("Hello World!")
            GetSingleSelection()
        End Sub
        Public Sub DrawSketchLine()
            ' Check to make sure a sketch is open.
            If Not TypeOf ThisApplication.ActiveEditObject Is PlanarSketch Then
                MsgBox("A sketch must be active.")
                Exit Sub
            End If

            ' Set a reference to the active sketch.
            Dim oSketch As PlanarSketch
            oSketch = ThisApplication.ActiveEditObject

            ' Set a reference to the transient geometry collection.
            Dim oTransGeom As TransientGeometry
            oTransGeom = ThisApplication.TransientGeometry

            ' Create a new transaction to wrap the construction of the three lines
            ' into a single undo.
            Dim oTrans As Transaction
            oTrans = ThisApplication.TransactionManager.StartTransaction( _
            ThisApplication.ActiveDocument, _
            "Create Triangle Sample")

            ' Create the first line of the triangle. This uses two transient points as
            ' input to definethe coordinates of the ends of the line. Since a transient
            ' point is input a sketch point is automatically created at that location and
            ' the line is attached to it.
            Dim oLines(0 To 2) As SketchLine
            oLines(0) = oSketch.SketchLines.AddByTwoPoints(oTransGeom.CreatePoint2d(0, 0), _
            oTransGeom.CreatePoint2d(4, 0))

            ' Create a sketch line that is connected to the sketch point the previous lines
            ' end point is connected to. This will automatically create the constraint to
            ' tie the new line to the sketch point the previous line is also connected to.
            ' This will result in the two lines being connected since they're both tied
            ' to the same sketch point.
            oLines(1) = oSketch.SketchLines.AddByTwoPoints(oLines(0).EndSketchPoint, _
            oTransGeom.CreatePoint2d(2, 3))

            ' Create a third line and connect it to the start point of the first line and the
            ' end point of the second line. This will result in a connected triangle.
            oLines(2) = oSketch.SketchLines.AddByTwoPoints(oLines(1).EndSketchPoint, _
            oLines(0).StartSketchPoint)

            ' End the transaction for the triangle.
            oTrans.End()

            ' Create a rectangle whose lines are parallel to the sketch planes x and y axes
            ' by using the SketchLines.AddAsTwoPointRectangle method. The top point of the
            ' triangle will be used as input for one of the points so the rectangle will be
            ' tied to that point.
            Dim oRectangleLines As SketchEntitiesEnumerator
            oRectangleLines = oSketch.SketchLines.AddAsTwoPointRectangle( _
            oLines(2).StartSketchPoint, _
            oTransGeom.CreatePoint2d(5, 5))

            ' Create a rotated rectangle by using the SketchLines.AddAsTwoRectangle method.
            ' One of the corners of this rectangle will be tied to the corner of the previous
            ' rectangle.
            oRectangleLines = oSketch.SketchLines.AddAsThreePointRectangle( _
            oRectangleLines(2).EndSketchPoint, _
            oTransGeom.CreatePoint2d(7, 3), _
            oTransGeom.CreatePoint2d(8, 8))
        End Sub
        ''' <summary>
        ''' Returns the Object Name from a picked Sketch
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub GetSingleSelection()
            ' Check to make sure a sketch is open.
            If Not TypeOf ThisApplication.ActiveEditObject Is PlanarSketch Then
                MsgBox("A sketch must be active.")
                Exit Sub
            End If
            m_activeSketch = ThisApplication.ActiveEditObject
            ' Get a feature selection from the user
            'Dim ske As SketchEntity
            m_skProfile = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kSketchProfileFilter, "Pick a Profile feature")
            'ske = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAllEntitiesFilter, "Pick the boundary")
            'ske = DirectCast(m_skProfile, SketchEntity)
            'If we can select a profile from the sketch we can simply call the profile.regionproperties.centroid method!
            Dim Centroid As Point2d = DirectCast(m_skProfile.RegionProperties.Centroid, Point2d)
            'we can take centroid and use it as a starting point for our Hyperbolic Tessellation(s)!
            Dim tr As Transaction = ThisApplication.TransactionManager.StartTransaction(ThisApplication.ActiveDocument, "IV_Tessellator")
            Try
                Dim minValue As Point2d
                Dim maxValue As Point2d
                'GetProfileRegionProperties(m_skProfile)
                Dim tmpBox As Box2d = ProfileRange(m_skProfile)
                minValue = tmpBox.MinPoint
                maxValue = tmpBox.MaxPoint
                createHyperbolicTilesIV(tr, Centroid, minValue, maxValue)
            Catch ex As System.Exception
                MessageBox.Show(ex.Message)
                tr.Abort()
            Finally
                tr.End()
            End Try
        End Sub



        Public Sub createHyperbolicTilesIV(trans As Transaction, cen As Point2d, min As Point2d, max As Point2d)
            'since we already picked an item that has a centre point, we can 
            ' ignore all the Prompts concerned with that!

            Dim n As Integer, k As Integer
            Dim done As Boolean

            Do
                Dim sides As Integer = InputBox("Number of sides for polygons", "Input a number < 20 & > 3!", "7")
                If sides <= 20 And sides >= 3 Then
                    n = sides
                End If
                Dim polysmeet As Integer = InputBox("Number of polygons meeting at each vertex", "Input a number > 2", "3")
                If polysmeet <= 5 And polysmeet >= 3 Then
                    k = polysmeet
                End If
                done = ((n - 2) * (k - 2) > 4)
                If Not done Then
                    MessageBox.Show("Please increase values for a successful tiling.")
                End If
            Loop While Not done

            Dim lvls As Integer = InputBox("Number of levels to generate?", "Levels to Generate", "3")
            If lvls > 5 Then
                MessageBox.Show("picking a number > 5 will cause uneccessary levels of detail!")
                Do While lvls > 5
                    lvls = InputBox("Number of levels to generate?", "Levels to Generate", "3")
                Loop
            End If
            Dim straight As Boolean = IIf(InputBox("Use curved or straight segments", DefaultResponse:="Straight") = "Straight", True, False)

            'Kean's code creates layers for the geometry to be created upon; Since we're already in an 
            ' inventor sketch, we don't need this!?
            Dim he As HyperbolicEngine = New HyperbolicEngine(trans, m_activeSketch, min, max, cen, straight)
            ' This entire method could be deprecated in favour of Kean's Windows-Azure based solution; all it would need is
            ' a httpwebrequest to return a json-encoded list of 3D points.
            Dim pe As PolygonEngine = New PolygonEngine()

            Dim polys As Polygon() = pe.GetPolygons(n, k, lvls)
            'need to figure out what the command is to cancel a command mid-flow;
            ' the AutoCAD equivalent is:
            'If HostApplicationServices.Current.UserBreak() Then
            '    Return
            'End If
            'something else from AutoCAD that doesn't work here and won't convert to something that should!
            'ed.WriteMessage(vbLf & "[{0} {1}] (level {2}) => {3} polygons.", n, k, lev, polys.Length)
            Dim i As Integer = 1
            For Each p As Polygon In polys
                If Not p Is Nothing Then
                    he.DrawPolygon(p.Vertices.ToArray())
                    i += 1
                    p.Vertices.Clear()
                End If
            Next

        End Sub

        Private Sub m_LinesButtonDef_OnExecute(Context As Inventor.NameValueMap) Handles m_LinesButtonDef.OnExecute
            If Not TypeOf ThisApplication.ActiveEditObject Is PlanarSketch Then
                MsgBox("A sketch must be active.")
                Exit Sub
            End If
            m_activeSketch = ThisApplication.ActiveEditObject
            ' Get a feature selection from the user
            ' Set a reference to the transient geometry collection.
            Dim oTransGeom As TransientGeometry
            oTransGeom = ThisApplication.TransientGeometry

            ' Create a new transaction to wrap the construction of the three lines
            ' into a single undo.
            Dim oTrans As Transaction
            oTrans = ThisApplication.TransactionManager.StartTransaction( _
            ThisApplication.ActiveDocument, _
            "Create Triangle Sample")

            ' Create the first line of the triangle. This uses two transient points as
            ' input to definethe coordinates of the ends of the line. Since a transient
            ' point is input a sketch point is automatically created at that location and
            ' the line is attached to it.
            Dim oLines(0 To 2) As SketchLine
            oLines(0) = m_activeSketch.SketchLines.AddByTwoPoints(oTransGeom.CreatePoint2d(50.5, 25.5), _
            oTransGeom.CreatePoint2d(90.5, 25.5))

            ' Create a sketch line that is connected to the sketch point the previous lines
            ' end point is connected to. This will automatically create the constraint to
            ' tie the new line to the sketch point the previous line is also connected to.
            ' This will result in the two lines being connected since they're both tied
            ' to the same sketch point.
            oLines(1) = m_activeSketch.SketchLines.AddByTwoPoints(oLines(0).EndSketchPoint, _
            oTransGeom.CreatePoint2d(70, 80))

            ' Create a third line and connect it to the start point of the first line and the
            ' end point of the second line. This will result in a connected triangle.
            oLines(2) = m_activeSketch.SketchLines.AddByTwoPoints(oLines(1).EndSketchPoint, _
            oLines(0).StartSketchPoint)

            ' End the transaction for the triangle.
            oTrans.End()

            ' Create a rectangle whose lines are parallel to the sketch planes x and y axes
            ' by using the SketchLines.AddAsTwoPointRectangle method. The top point of the
            ' triangle will be used as input for one of the points so the rectangle will be
            ' tied to that point.
            Dim oRectangleLines As SketchEntitiesEnumerator
            oRectangleLines = m_activeSketch.SketchLines.AddAsTwoPointRectangle( _
            oLines(2).StartSketchPoint, _
            oTransGeom.CreatePoint2d(5, 5))

            ' Create a rotated rectangle by using the SketchLines.AddAsTwoRectangle method.
            ' One of the corners of this rectangle will be tied to the corner of the previous
            ' rectangle.
            oRectangleLines = m_activeSketch.SketchLines.AddAsThreePointRectangle( _
            oRectangleLines(2).EndSketchPoint, _
            oTransGeom.CreatePoint2d(7, 3), _
            oTransGeom.CreatePoint2d(8, 8))


        End Sub

        Private Function ProfileRange(profile As Profile) As Box2d
            Dim oRange As Box2d
            oRange = profile.Item(1).Item(1).Curve.Evaluator.RangeBox
            'had to create a TransientGeometry Object as for some reason the oRange object doesn't update correctly!
            Dim tmprange As Box2d = ThisApplication.TransientGeometry.CreateBox2d
            Dim tmpMin As Point2d = ThisApplication.TransientGeometry.CreatePoint2d(0, 0)
            Dim tmpMax As Point2d = ThisApplication.TransientGeometry.CreatePoint2d(0, 0)
            Dim oPath As ProfilePath
            For Each oPath In profile
                Dim oEntity As ProfileEntity
                For Each oEntity In oPath
                    Dim oEval As Curve2dEvaluator
                    oEval = oEntity.Curve.Evaluator

                    Dim oCurveRange As Box2d
                    oCurveRange = oEval.RangeBox
                    ' Expand the current range box, if needed.
                    If oCurveRange.MinPoint.X < oRange.MinPoint.X Then
                        tmpMin.X = oCurveRange.MinPoint.X
                    Else
                        tmpMin.X = oRange.MinPoint.X
                    End If

                    If oCurveRange.MinPoint.Y < oRange.MinPoint.Y Then
                        tmpMin.Y = oCurveRange.MinPoint.Y
                    Else
                        tmpMin.Y = oRange.MinPoint.Y
                    End If

                    If oCurveRange.MaxPoint.X > oRange.MaxPoint.X Then
                        tmpMax.X = oCurveRange.MaxPoint.X
                    Else
                        tmpMax.X = oRange.MaxPoint.X
                    End If

                    If oCurveRange.MaxPoint.Y > oRange.MaxPoint.Y Then
                        tmpMax.Y = oCurveRange.MaxPoint.Y
                    Else
                        tmpMax.Y = oRange.MaxPoint.Y
                    End If
                Next
            Next
            tmprange.MinPoint = tmpMin
            tmprange.MaxPoint = tmpMax
            Return tmprange
        End Function

        'Public Sub GetProfileRegionProperties(profile As Profile)
        '    ' Set a reference to the region properties object.
        '    Dim oRegionProps As RegionProperties
        '    oRegionProps = profile.RegionProperties

        '    ' Set the accuracy to medium.
        '    oRegionProps.Accuracy = Inventor.AccuracyEnum.kMedium

        '    ' Display the region properties of the profile.
        '    Debug.Print("Area: " & oRegionProps.Area)

        '    Debug.Print("Perimeter: " & oRegionProps.Perimeter)

        '    Debug.Print("Centroid: " & oRegionProps.Centroid.X & ", " & oRegionProps.Centroid.Y)

        '    Dim adPrincipalMoments(0 To 2) As Double
        '    oRegionProps.PrincipalMomentsOfInertia(adPrincipalMoments(0), adPrincipalMoments(1), adPrincipalMoments(2))
        '    Debug.Print("Principal Moments of Inertia: " & adPrincipalMoments(0) & ", " & adPrincipalMoments(1))

        '    Dim adRadiusOfGyration(0 To 2) As Double
        '    oRegionProps.RadiusOfGyration(adRadiusOfGyration(0), adRadiusOfGyration(1), adRadiusOfGyration(2))
        '    Debug.Print("Radius of Gyration: " & adRadiusOfGyration(0) & ", " & adRadiusOfGyration(1))

        '    Dim Ixx As Double
        '    Dim Iyy As Double
        '    Dim Izz As Double
        '    Dim Ixy As Double
        '    Dim Iyz As Double
        '    Dim Ixz As Double
        '    oRegionProps.MomentsOfInertia(Ixx, Iyy, Izz, Ixy, Iyz, Ixz)
        '    Debug.Print("Moments: ")
        '    Debug.Print(" Ixx: " & Ixx)
        '    Debug.Print(" Iyy: " & Iyy)
        '    Debug.Print(" Ixy: " & Ixy)

        '    Dim xVector As Vector = Nothing
        '    Dim yVector As Vector = Nothing
        '    oRegionProps.PrincipalAxes(xVector, yVector)
        '    Debug.Print("Axes: ")
        '    Debug.Print("X Axes: " & xVector.ToString)
        '    Debug.Print("Y Axes: " & yVector.ToString)

        '    Debug.Print("Rotation Angle from projected Sketch Origin to Principle Axes: " _
        '    & oRegionProps.RotationAngle)

        'End Sub


        ''' <summary>
        ''' The object created does not represent a coordinate system but instead is a representation 
        ''' of the information that defines a coordinate system. You can use this object as input to the 
        ''' UserCoordinateSystems.Add method to create the actual coordinate system. 
        ''' The returned UserCoordinateSystemDefinition object returned is fully defined (initialized to an identity matrix)
        ''' and can be used to create a coordinate system. However, you may want to change some of the values of the 
        ''' UserCoordinateSystemDefinition object before using it to create a coordinate system.
        ''' </summary>
        ''' <remarks>Copied from the Inventor 2012 COM API Reference file!</remarks>
        Public Sub CreateUCSByTransformationMatrix()

            ' Create a new part document
            Dim oDoc As PartDocument
            oDoc = ThisApplication.Documents.Add(DocumentTypeEnum.kPartDocumentObject)

            ' Set a reference to the PartComponentDefinition object
            Dim oCompDef As PartComponentDefinition
            oCompDef = oDoc.ComponentDefinition

            Dim oTG As TransientGeometry
            oTG = ThisApplication.TransientGeometry

            ' Create an identity matrix
            Dim oMatrix As Matrix
            oMatrix = oTG.CreateMatrix

            ' Rotate about Z-Axis by 45 degrees
            oMatrix.SetToRotation(3.14159 / 4, oTG.CreateVector(0, 0, 1), oTG.CreatePoint(0, 0, 0))

            ' Translate the origin to (2, 2, 2)
            Dim oTranslationMatrix As Matrix
            oTranslationMatrix = oTG.CreateMatrix
            oTranslationMatrix.SetTranslation(oTG.CreateVector(2, 2, 2))

            oMatrix.TransformBy(oTranslationMatrix)

            ' Create an empty definition object
            Dim oUCSDef As UserCoordinateSystemDefinition
            oUCSDef = oCompDef.UserCoordinateSystems.CreateDefinition

            ' Set it to be based on the defined matrix
            oUCSDef.Transformation = oMatrix

            ' Create the UCS
            Dim oUCS As UserCoordinateSystem
            oUCS = oCompDef.UserCoordinateSystems.Add(oUCSDef)

        End Sub
#End Region
    End Class
#Region "Worker Classes"
    Public Class HyperbolicEngine
        'get the inventor instance
        Private thisApp As Inventor.Application = StandardAddInServer.ThisApplication
        Private trans As Transaction
        Private sketchEnt As SketchEntity

        ' We'll also hold information about the target circle

        Private radius As Double
        Private min As Double
        Private max As Double
        Private planarSketch As PlanarSketch
        Private center As Inventor.Point2d
        Private displacement As Matrix2d

        ' Whether we'll draw curved or straight lines

        Private straight As Boolean

        ' Target layers for points, supporting lines and our
        ' actual geometry

        'Private _ptLay As Long, _lnLay As Long, _hgLay As Long

        ' Constructor

        Public Sub New(tr As Transaction, activeSketch As PlanarSketch, min As Point2d, max As Point2d, center As Point2d, straight As Boolean)
            trans = tr
            planarSketch = activeSketch
            center = center
            radius = DistPointPoint(min, max)
            displacement = thisApp.TransientGeometry.CreateMatrix2d()
            displacement.SetTranslation(thisApp.TransientGeometry.CreateVector2d(center.X, center.Y))
            straight = straight
        End Sub

        ' Functions to map a point from the unit disk to WCS

        Private Function WcsPoint(x As Double, y As Double) As Point2d
            Dim pt As Point2d = thisApp.TransientGeometry.CreatePoint2d(x, y)
            'the next line should work but doesn't! So we'll skip it for the time being!
            pt.TranslateBy(displacement.Translation)
            'pt.TranslateBy(_disp)
            Return pt
        End Function
        Private Function multiplyByRadius(c As Double) As Double
            Return c * radius / 2 'multiply by the radius cubed to see if working too near the 0,0 origin is an issue
        End Function

        Private Function cX(x As Double) As Double
            Return x * radius
        End Function

        Private Function cY(y As Double) As Double
            Return y * radius
        End Function

        ''' <summary>
        ''' Compute the circle containing points a and b, and
        ''' orthogonal to the unit circle (a.k.a shortest path between
        ''' a and b in the hyperbolic sense).
        ''' Return true if the output is a circle. False if the shortest path is a straight line (diameter of the unit circle).
        '''</summary>
        ''' <param name="a"></param>
        ''' <param name="b"></param>
        ''' <param name="cx"></param>
        ''' <param name="cy"></param>
        ''' <param name="radius"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function ComputeCircleParameters(a As Point2d, b As Point2d, ByRef cx As Double, ByRef cy As Double, ByRef radius As Double) As Boolean
            Dim ax As Double = a.X, ay As Double = a.Y, bx As Double = b.X, by As Double = b.Y
            Dim tol As Double = 0.0005

            cx = 0
            cy = 0
            radius = 0

            If Math.Abs(ax * by - ay * bx) < tol Then
                Return False
            End If

            cx = (-1 / 2.0 * (ay * bx * bx + ay * by * by - (ax * ax + ay * ay + 1) * by + ay) / (ax * by - ay * bx))

            cy = (1 / 2.0 * (ax * bx * bx + ax * by * by - (ax * ax + ay * ay + 1) * bx + ax) / (ax * by - ay * bx))

            radius = -1 / 2.0 * Math.Sqrt(ax * ax * ax * ax * bx * bx + ax * ax * ax * ax * by * by - 2 * ax * ax * ax * bx * bx * bx - 2 * ax * ax * ax * bx * by * by + 2 * ax * ax * ay * ay * bx * bx + 2 * ax * ax * ay * ay * by * by - 2 * ax * ax * ay * bx * bx * by - 2 * ax * ax * ay * by * by * by + ax * ax * bx * bx * bx * bx + 2 * ax * ax * bx * bx * by * by + ax * ax * by * by * by * by - 2 * ax * ay * ay * bx * bx * bx - 2 * ax * ay * ay * bx * by * by + ay * ay * ay * ay * bx * bx + ay * ay * ay * ay * by * by - 2 * ay * ay * ay * bx * bx * by - 2 * ay * ay * ay * by * by * by + ay * ay * bx * bx * bx * bx + 2 * ay * ay * bx * bx * by * by + ay * ay * by * by * by * by - 2 * ax * ax * ax * bx - 2 * ax * ax * ay * by + 4 * ax * ax * bx * bx - 2 * ax * ay * ay * bx + 8 * ax * ay * bx * by - 2 * ax * bx * bx * bx - 2 * ax * bx * by * by - 2 * ay * ay * ay * by + 4 * ay * ay * by * by - 2 * ay * bx * bx * by - 2 * ay * by * by * by + ax * ax - 2 * ax * bx + ay * ay - 2 * ay * by + bx * bx + by * by) / (ax * by - ay * bx)

            If radius < 0.0 Then
                radius = 1 / 2.0 * Math.Sqrt(ax * ax * ax * ax * bx * bx + ax * ax * ax * ax * by * by - 2 * ax * ax * ax * bx * bx * bx - 2 * ax * ax * ax * bx * by * by + 2 * ax * ax * ay * ay * bx * bx + 2 * ax * ax * ay * ay * by * by - 2 * ax * ax * ay * bx * bx * by - 2 * ax * ax * ay * by * by * by + ax * ax * bx * bx * bx * bx + 2 * ax * ax * bx * bx * by * by + ax * ax * by * by * by * by - 2 * ax * ay * ay * bx * bx * bx - 2 * ax * ay * ay * bx * by * by + ay * ay * ay * ay * bx * bx + ay * ay * ay * ay * by * by - 2 * ay * ay * ay * bx * bx * by - 2 * ay * ay * ay * by * by * by + ay * ay * bx * bx * bx * bx + 2 * ay * ay * bx * bx * by * by + ay * ay * by * by * by * by - 2 * ax * ax * ax * bx - 2 * ax * ax * ay * by + 4 * ax * ax * bx * bx - 2 * ax * ay * ay * bx + 8 * ax * ay * bx * by - 2 * ax * bx * bx * bx - 2 * ax * bx * by * by - 2 * ay * ay * ay * by + 4 * ay * ay * by * by - 2 * ay * bx * bx * by - 2 * ay * by * by * by + ax * ax - 2 * ax * bx + ay * ay - 2 * ay * by + bx * bx + by * by) / (ax * by - ay * bx)
            End If

            Return True
        End Function

        ' Draw the unit circle

        Public Sub DrawUnitCircle()
            If trans IsNot Nothing Then
                Dim oCircle As SketchCircle
                oCircle = planarSketch.SketchCircles.AddByCenterRadius(center, radius)
                'Dim c As New Circle(_center, Vector3d.ZAxis, _radius)
                'c.LayerId = _hgLay
                '_tr.AddNewlyCreatedDBObject(oCircle, True)
            End If
        End Sub

        ' Draw a point in the Poincare disc

        Public Sub DrawPoint(a As Point2d)
            If trans IsNot Nothing Then
                Dim oPoint As Point2d = thisApp.TransientGeometry.CreatePoint2d(a.X, a.Y)
                Dim sPoint As SketchPoint
                sPoint = planarSketch.SketchPoints.Add(oPoint)
                'Dim p As New DBPoint(WcsPoint(a.X, a.Y))
                'p.LayerId = _ptLay
                '_tr.AddNewlyCreatedDBObject(p, True)
            End If
        End Sub

        ' Draw a hyperbolic line through a and b

        Public Sub DrawSupportingLine(a As Point2d, b As Point2d, withPoints As Boolean)
            If withPoints Then
                DrawPoint(a)
                DrawPoint(b)
            End If

            Dim ax As Double = a.X, ay As Double = a.Y, bx As Double = b.X, by As Double = b.Y
            Dim cx As Double, cy As Double, r As Double

            Dim result As Boolean = ComputeCircleParameters(a, b, cx, cy, r)

            'Dim ent As Entity

            Dim newLine As SketchLine
            Dim newCircle As SketchCircle


            If Not result Then
                ' We project a and b points onto the unit circle

                Dim theta As Double = Math.Atan2(ay, ax)
                Dim theta2 As Double = Math.Atan2(by, bx)
                newLine = planarSketch.SketchLines.AddByTwoPoints(WcsPoint(Math.Cos(theta), Math.Sin(theta)), WcsPoint(Math.Cos(theta2), Math.Sin(theta2)))
                'ent = New Line(WcsPoint(Math.Cos(theta), Math.Sin(theta)), WcsPoint(Math.Cos(theta2), Math.Sin(theta2)))
            Else
                newCircle = planarSketch.SketchCircles.AddByCenterRadius(WcsPoint(cx, cy), r * radius)
                'ent = New Circle(WcsPoint(cx, cy), Vector3d.ZAxis, r * _radius)
            End If
        End Sub

        ' Draw a hyperbolic edge between a and b

        Public Sub DrawEdge(a As Point2d, b As Point2d, withPoints As Boolean, Optional straightSegments As Boolean = False)
            If withPoints Then
                DrawPoint(a)
                DrawPoint(b)
            End If

            Dim ax As Double = a.X, ay As Double = a.Y, bx As Double = b.X, by As Double = b.Y
            Dim cx As Double = 0.0, cy As Double = 0.0, r As Double = 0.0
            Dim result As Boolean = If(straight, False, ComputeCircleParameters(a, b, cx, cy, r))

            ' Near-aligned points -> straight segment

            'Dim ent As Entity = Nothing
            Dim newArc As SketchArc
            Dim newLine As SketchLine

            If Not result Then
                newLine = planarSketch.SketchLines.AddByTwoPoints(WcsPoint(ax, ay), WcsPoint(bx, by)) ' this line fails when it reaches the last point?
                'newLine = New Line(WcsPoint(ax, ay), WcsPoint(bx, by))
            Else
                Dim theta As Double = -Math.Atan2(-ay + cy, ax - cx)
                Dim theta2 As Double = -Math.Atan2(-by + cy, bx - cx)

                ' We recenter the angles to [0, 2Pi]

                While theta < 0
                    theta += 2 * Math.PI
                End While
                While theta2 < 0
                    theta2 += 2 * Math.PI
                End While

                Dim cen As Point2d = WcsPoint(cx, cy)

                Dim delta As Double = Math.Abs(theta - theta2)

                If (theta > theta2 AndAlso delta > Math.PI) OrElse (theta < theta2 AndAlso delta < Math.PI) Then
                    newArc = planarSketch.SketchArcs.AddByCenterStartEndPoint(cen, theta, theta2)
                    'ent = New Arc(cen, r * _radius, theta, theta2)
                Else
                    newArc = planarSketch.SketchArcs.AddByCenterStartEndPoint(cen, theta, theta2)
                    'ent = New Arc(cen, r * _radius, theta2, theta)
                End If
            End If
        End Sub

        ' Draw a hyperbolic triangle with vertices (a,b,c)

        Public Sub DrawTriangle(VerticeA As Point2d, VerticeB As Point2d, VerticeC As Point2d, Optional withPoints As Boolean = False, Optional withSupportLines As Boolean = False)
            DrawPolygon(New Point2d() {VerticeA, VerticeB, VerticeC}, withPoints, withSupportLines)
        End Sub

        ' Draw a hyperbolic polygon with the provided vertices

        Public Sub DrawPolygon(pts As Point2d(), Optional withPoints As Boolean = False, Optional withSupportLines As Boolean = False)
            Dim oTransGeom As TransientGeometry = thisApp.TransientGeometry
            If withPoints Then
                For Each pt As Point2d In pts
                    DrawPoint(pt)
                Next
            End If

            Dim numPts As Integer = pts.Length

            If withSupportLines AndAlso Not straight Then
                For i As Integer = 0 To numPts - 1
                    DrawSupportingLine(pts(i), pts((i + 1) Mod numPts), withPoints)
                Next
            End If

            Dim oLines(0 To numPts - 1) As SketchLine

            For i As Integer = 0 To numPts - 1
                Dim ax As Double = pts(i).X
                Dim ay As Double = pts(i).Y
                Dim aPoint As Point2d = oTransGeom.CreatePoint2d(ax, ay)
                aPoint.X = multiplyByRadius(aPoint.X)
                aPoint.Y = multiplyByRadius(aPoint.Y)
                Dim bPoint As Point2d = pts((i + 1) Mod numPts)
                bPoint.X = multiplyByRadius(bPoint.X)
                bPoint.Y = multiplyByRadius(bPoint.Y)
                'translate Points a & b to the correct location!?
                aPoint = WcsPoint(aPoint.X, aPoint.Y)
                bPoint = WcsPoint(bPoint.X, bPoint.Y)
                Dim cx As Double = 0.0, cy As Double = 0.0, r As Double = 0.0
                Dim result As Boolean = If(straight, False, ComputeCircleParameters(aPoint, bPoint, cx, cy, r))
                Select Case i
                    Case 0 'polygon start
                        oLines(i) = planarSketch.SketchLines.AddByTwoPoints(oTransGeom.CreatePoint2d(aPoint.X, aPoint.Y), oTransGeom.CreatePoint2d(bPoint.X, bPoint.Y))
                    Case numPts - 1 'polygon end
                        oLines(i) = planarSketch.SketchLines.AddByTwoPoints(oLines(i - 1).EndSketchPoint, oLines(0).StartSketchPoint)
                    Case Else ' add the end of the last line as the start of this one.
                        oLines(i) = planarSketch.SketchLines.AddByTwoPoints(oLines(i - 1).EndSketchPoint, _
                                                                        oTransGeom.CreatePoint2d(bPoint.X, bPoint.Y))
                End Select
                'DrawEdge(pts(i), pts((i + 1) Mod numPts), withPoints)
            Next
        End Sub

        Private Function DistPointPoint(point2d As Point2d, point2d1 As Point2d) As Double
            'compute the distance between the minpoint and maxpoints of the object.
            Dim dDist As Double
            Dim lSegment As LineSegment2d = thisApp.TransientGeometry.CreateLineSegment2d(point2d, point2d1)
            dDist = lSegment.StartPoint.DistanceTo(lSegment.EndPoint)
            'dDist = lSegment.DistanceTo(point2d1)
            ' Create a vector that defines the direction from the lines origin to the input point.
            'Dim oPointVec As Vector = Line.RootPoint.VectorTo(Point)

            ' Get the angle between the two vectors.
            'Dim oAngle As Double
            'oAngle() = Line.Direction.AngleTo(oPointVec.AsUnitVector)
            ' Calculate the distance using law of sines.
            'DistPointLine = dDist / Sin(oAngle)
            Return dDist
        End Function

    End Class

    Public Class Polygon
        'Public Vertices As Collection()
        Public Vertices As New List(Of Point2d)
    End Class



    Public Class PolygonEngine
        ' get the Inventor instance
        Private thisApp As Inventor.Application = StandardAddInServer.ThisApplication
        ' The list of polygons

        ' the active sketch
        Private activeSketch As PlanarSketch = thisApp.ActiveEditObject

        Private polys As Polygon()

        ' Previously created neighbors for the polygons

        Private rules As Integer()

        ' The total number of polygons in all the layers

        Private totalPolygons As Integer

        ' The number through one less layer

        Private innerPolygons As Integer

        Private tol As Double = 0.0005
        Public Function GetPolygons(n As Integer, k As Integer, layNum As Integer) As Polygon()
            CountPolygons(layNum, n, k)
            DeterminePolygons(n, k)

            Return polys
        End Function

        Private Function ConstructCenterPolygon(n As Integer, k As Integer) As Polygon
            ' Initialize P as the center polygon in an n-k regular tiling
            ' Let ABC be a triangle in a regular (n,k0-tiling, where
            '    A is the center of an n-gon (also center of the disk),
            '    B is a vertex of the n-gon, and
            '    C is the midpoint of a side of the n-gon adjacent to B

            Dim angleA As Double = Math.PI / n
            Dim angleB As Double = Math.PI / k
            Dim angleC As Double = Math.PI / 2.0

            ' For a regular tiling, we need to compute the distance s
            ' from A to B

            Dim sinA As Double = Math.Sin(angleA)
            Dim sinB As Double = Math.Sin(angleB)
            Dim s As Double = Math.Sin(angleC - angleB - angleA) / Math.Sqrt(1.0 - sinB * sinB - sinA * sinA)

            ' Determine the coordinates of the n vertices of the n-gon.
            ' They're all at distance s from center of the Poincare disk

            Dim p As New Polygon()

            For i As Integer = 0 To n - 1
                Dim pt As Point2d = thisApp.TransientGeometry.CreatePoint2d(s * Math.Cos((3 + 2 * i) * angleA), s * Math.Sin((3 + 2 * i) * angleA))
                p.Vertices.Add(pt)

            Next
            Return p
        End Function

        Private Sub CountPolygons(layer As Integer, n As Integer, k As Integer)
            ' Determine
            '   totalPolygons:  the number of polygons there are through
            '                   that many layers
            '   innerPolygons:  the number through one less layer

            totalPolygons = 1
            ' count the central polygon
            innerPolygons = 0

            Dim a As Integer = n * (k - 3)
            ' polygons in 1st layer joined by vertex
            Dim b As Integer = n
            ' polygons in 1st layer joined by edge
            Dim next_a As Integer, next_b As Integer
            For l As Integer = 1 To layer
                innerPolygons = totalPolygons
                If k = 3 Then
                    next_a = a + b
                    next_b = (n - 6) * a + (n - 5) * b
                Else
                    ' k > 3
                    next_a = ((n - 2) * (k - 3) - 1) * a + ((n - 3) * (k - 3) - 1) * b
                    next_b = (n - 2) * a + (n - 3) * b
                End If
                totalPolygons += a + b
                a = next_a
                b = next_b
            Next
        End Sub

        ' Rule codes
        '     *   0:  initial polygon.  Needs neighbors on all n sides
        '     *   1:  polygon already has 2 neighbors, but one less
        '     *       around corner needed
        '     *   2:  polygon already has 2 neighbors
        '     *   3:  polygon already has 3 neighbors
        '     *   4:  polygon already has 4 neighbors
        '     

        Private Sub DeterminePolygons(n As Integer, k As Integer)
            polys = New Polygon(totalPolygons) {}
            rules = New Integer(totalPolygons) {}
            polys(0) = ConstructCenterPolygon(n, k)
            rules(0) = 0
            Dim j As Integer = 1
            ' index of the next polygon to create
            For i As Integer = 0 To innerPolygons - 1
                j = ApplyRule(i, j, n, k)
            Next
        End Sub

        Private Function ApplyRule(i As Integer, j As Integer, n As Integer, k As Integer) As Integer
            Dim r As Integer = rules(i)
            Dim special As Boolean = (r = 1)
            If special Then
                r = 2
            End If
            Dim start As Integer = If((r = 4), 3, 2)
            Dim quantity As Integer = If((k = 3 AndAlso r <> 0), n - r - 1, n - r)
            For s As Integer = start To start + (quantity - 1)
                ' Create a polygon adjacent to P[i]

                polys(j) = CreateNextPolygon(polys(i), s Mod n, n, k)
                rules(j) = If((k = 3 AndAlso s = start AndAlso r <> 0), 4, 3)
                j += 1
                Dim m As Integer
                If special Then
                    m = 2
                ElseIf s = 2 AndAlso r <> 0 Then
                    m = 1
                Else
                    m = 0
                End If
                While m < k - 3
                    ' Create a polygon adjacent to P[j-1]

                    polys(j) = CreateNextPolygon(polys(j - 1), 1, n, k)
                    rules(j) = If((n = 3 AndAlso m = k - 4), 1, 2)
                    j += 1
                    m += 1
                End While
            Next
            Return j
        End Function

        ' Reflect P thru the point or the side indicated by the side s
        ' to produce the resulting polygon Q


        Private Function CreateNextPolygon(P As Polygon, s As Integer, n As Integer, k As Integer) As Polygon
            Dim start As Point2d = P.Vertices(s), [end] As Point2d = P.Vertices((s + 1) Mod n)

            Dim Q As New Polygon()
            For i As Integer = 0 To n - 1
                Dim newpoint As Point2d = thisApp.TransientGeometry.CreatePoint2d(activeSketch.OriginPointGeometry.X, activeSketch.OriginPointGeometry.Y)
                Q.Vertices.Add(newpoint)
                'Q.Vertices.Add(Point2d.Origin)
            Next

            For i As Integer = 0 To n - 1
                ' Reflect P[i] thru C to get Q[j]}
                Dim j As Integer = (n + s - i + 1) Mod n
                Q.Vertices(j) = Reflect(start, [end], P.Vertices(i))
            Next
            Return Q
        End Function

        ' Reflect the point R across the line through A to B

        Public Function Reflect(A As Point2d, B As Point2d, R__1 As Point2d) As Point2d
            ' Find a unit vector D in the direction of the line

            Dim den As Double = A.X * B.Y - B.X * A.Y
            Dim straight As Boolean = Math.Abs(den) < tol

            If straight Then
                ' The less interesting case - we could do this less
                ' manually using AcGe

                Dim P As Point2d = A
                den = Math.Sqrt((A.X - B.X) * (A.X - B.X) + (A.Y - B.Y) * (A.Y - B.Y))
                Dim D As Point2d = thisApp.TransientGeometry.CreatePoint2d((B.X - A.X) / den, (B.Y - A.Y) / den)

                ' Reflect method

                Dim factor As Double = 2.0 * ((R__1.X - P.X) * D.X + (R__1.Y - P.Y) * D.Y)
                Dim newPoint As Point2d = thisApp.TransientGeometry.CreatePoint2d(2.0 * P.X + factor * D.X - R__1.X, 2.0 * P.Y + factor * D.Y - R__1.Y)
                Return newPoint
            Else
                ' Find the center of the circle through these points
                ' (this is what we really need to use this for)

                Dim s1 As Double = (1.0 + A.X * A.X + A.Y * A.Y) / 2.0
                Dim s2 As Double = (1.0 + B.X * B.X + B.Y * B.Y) / 2.0

                Dim C As Point2d = thisApp.TransientGeometry.CreatePoint2d((s1 * B.Y - s2 * A.Y) / den, (A.X * s2 - B.X * s1) / den)
                Dim r__2 As Double = Math.Sqrt(C.X * C.X + C.Y * C.Y - 1.0)

                ' Reflect method

                Dim factor As Double = r__2 * r__2 / ((R__1.X - C.X) * (R__1.X - C.X) + (R__1.Y - C.Y) * (R__1.Y - C.Y))
                Dim newPoint As Point2d = thisApp.TransientGeometry.CreatePoint2d(C.X + factor * (R__1.X - C.X), C.Y + factor * (R__1.Y - C.Y))
                Return newPoint
            End If
        End Function
    End Class
#End Region
End Namespace

