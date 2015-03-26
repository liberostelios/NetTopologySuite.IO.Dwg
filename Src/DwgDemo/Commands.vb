Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices

Public Class Commands

    <CommandMethod("ReadBlockFeatures")> _
    Public Sub ReadBlockFeatures()
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
        Dim PSO As New PromptStringOptions("Please give me the name of the block table")
        Dim Result As PromptResult = ed.GetString(PSO)
        If Result.Status = PromptStatus.OK Then
            Dim Features As NetTopologySuite.Features.FeatureCollection = ReadBlockGeometry(Result.StringResult)
            MsgBox("Number of features: " + Features.Count.ToString())
        End If
    End Sub

    <CommandMethod("PolygonizeLines")> _
    Public Sub PolygonizeLines()
        Dim Polygons As List(Of GeoAPI.Geometries.IGeometry) = ReadLineGeometry("*", True)
        MsgBox("Number of polygons: " + Polygons.Count.ToString())
    End Sub

    Public Function ReadBlockGeometry(BlockName As String) As NetTopologySuite.Features.FeatureCollection
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        Dim values() As TypedValue = {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.BlockName, BlockName)}
        Dim filter As New SelectionFilter(values)

        Dim result As PromptSelectionResult = ed.SelectAll(filter)
        If result.Status = PromptStatus.OK Then
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Dim tr As Transaction = db.TransactionManager.StartOpenCloseTransaction()
            Dim reader As New NetTopologySuite.IO.Dwg.DwgReader()
            Dim rows As New NetTopologySuite.Features.FeatureCollection()

            For i As Integer = 0 To result.Value.Count - 1
                Dim ent As Entity = tr.GetObject(result.Value(i).ObjectId, OpenMode.ForRead)
                rows.Add(reader.ReadBlockReference(ent))
            Next

            tr.Commit()
            tr.Dispose()

            Return rows
        End If

        Return New NetTopologySuite.Features.FeatureCollection()
    End Function

    Public Function ReadLineGeometry(LayerFilter As String, Optional Polygonize As Boolean = False) As List(Of GeoAPI.Geometries.IGeometry)
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        Dim values() As TypedValue = {New TypedValue(DxfCode.Start, "*LINE"), New TypedValue(DxfCode.LayerName, LayerFilter)}
        Dim filter As New SelectionFilter(values)

        Dim result As PromptSelectionResult = ed.SelectAll(filter)
        If result.Status = PromptStatus.OK Then
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Dim tr As Transaction = db.TransactionManager.StartOpenCloseTransaction()
            Dim reader As New NetTopologySuite.IO.Dwg.DwgReader()
            Dim lines As New List(Of GeoAPI.Geometries.IGeometry)()

            For i As Integer = 0 To result.Value.Count - 1
                Dim ent As Entity = tr.GetObject(result.Value(i).ObjectId, OpenMode.ForRead)
                Dim line As GeoAPI.Geometries.IGeometry = reader.ReadGeometry(ent)
                lines.Add(line)
            Next

            tr.Commit()
            tr.Dispose()

            If (Polygonize) Then
                Dim polygonizer As New NetTopologySuite.Operation.Polygonize.Polygonizer()
                polygonizer.Add(lines)
                Return polygonizer.GetPolygons()
            Else
                Return lines
            End If
        End If

        Return New List(Of GeoAPI.Geometries.IGeometry)()
    End Function
End Class
