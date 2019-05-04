
Sub Create_Slicer()
'Declaring Variables
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim sc As SlicerCache
    Dim sl As Slicer

'faster code
Application.ScreenUpdating = False

'set sheet,pivottable and cache
    Set ws = Worksheets("Quarterly Summaries")
    Set pt = ws.PivotTables("PivotTable6")
  
'Creating slicer cache for Region
    Set sc = ThisWorkbook.SlicerCaches.Add2( _
    pt, _
    "Region", _
    "RegionSlicerCache", _
    XlSlicerCacheType.xlSlicer)
'Creating visual slicer
    Set sl = sc.Slicers.Add(ws, , "RegionSlicer", "Regions", _
    35, 0, 144, 198.75)

'all connect pivottable to slicer
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable7")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable9")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable15")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable4")

'Creating slicer cache for Offer Quarter
    Set sc = ThisWorkbook.SlicerCaches.Add2( _
    pt, _
    "Offer Quarter", _
    "OfferQuarterSlicerCache", _
    XlSlicerCacheType.xlSlicer)
'Creating visual slicer
    Set sl = sc.Slicers.Add(ws, , "OfferQuarterSlicer", "Offer Quarter", _
    35, 150, 144, 198.75)

'all connect pivottable to slicer
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable7")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable9")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable15")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable4")

'Creating slicer cache for Offer Quarter
    Set sc = ThisWorkbook.SlicerCaches.Add2( _
    pt, _
    "Position Type", _
    "PositionTypeSlicerCache", _
    XlSlicerCacheType.xlSlicer)
'Creating visual slicer
    Set sl = sc.Slicers.Add(ws, , "PositionTypeSlicer", "Postion Type", _
    35, 300, 144, 198.75)

'all connect pivottable to slicer
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable7")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable9")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable15")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable4")
    
'Creating slicer cache for Offer Quarter
    Set sc = ThisWorkbook.SlicerCaches.Add2( _
    pt, _
    "Org 2 Manager", _
    "Org2ManagerSlicerCache", _
    XlSlicerCacheType.xlSlicer)
'Creating visual slicer
    Set sl = sc.Slicers.Add(ws, , "Org2ManagerSlicer", "Org 2 Manager", _
    35, 450, 144, 198.75)

'all connect pivottable to slicer
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable7")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable9")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable15")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable4")
    
'Creating slicer cache for Offer Quarter
    Set sc = ThisWorkbook.SlicerCaches.Add2( _
    pt, _
    "Is Deep Learning Job?", _
    "DeeplearningSlicerCache", _
    XlSlicerCacheType.xlSlicer)
'Creating visual slicer
    Set sl = sc.Slicers.Add(ws, , "DeeplearningjobSlicer", "Deep Learning", _
    35, 600, 144, 198.75)

'all connect pivottable to slicer
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable7")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable9")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable15")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable4")

'Creating slicer cache for Offer Quarter
    Set sc = ThisWorkbook.SlicerCaches.Add2( _
    pt, _
    "Career Level Category", _
    "CareerBandSlicerCache", _
    XlSlicerCacheType.xlSlicer)
'Creating visual slicer
    Set sl = sc.Slicers.Add(ws, , "CareerBandjobSlicer", "Career Band", _
    35, 750, 144, 198.75)

'all connect pivottable to slicer
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable7")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable9")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable15")
    sc.PivotTables.AddPivotTable ws.PivotTables("PivotTable4")

'faster
Application.ScreenUpdating = True

End Sub
