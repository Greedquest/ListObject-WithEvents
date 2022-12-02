# ListObject-WithEvents

Get events when your Excel Tables (ListObjects) expand to hold new data, values change, rows are added and deleted and more...

This repo follows Test Driven Development. As such all of its features are self documenting:

 - [Default Categoriser tests](src/Watcher_DefaultCat_tests.bas)
 - [_NoCache_ Categoriser tests](src/Watcher_NoCacheCat_tests.bas)
 
## Quickstart
```vba
Private WithEvents fooTableEvents As TableWatcher

Sub StartListening()
   Set myTable = TableWatcher.Create(Sheet1.ListObjects("foo"))
End Sub

Private Sub fooTableEvents_RowAppended(ByVal where As ListRow)
  Debug.Print "New Row added to table Foo -"; where.DataBodyRange.Address
End Sub
```

## Cache vs NoCache
The constructor optionally takes a categoriser  - an object which converts `Worksheet_Change` events into `TableWatcher` events.
```vba
TableWatcher.Create(ByVal srcTable As ListObject, Optional ByVal eventsCategoriser As IWorksheetChangeCategoriser)
```
2 categorisers are provided to get you started:
 1. [DefaultCategoriser](src/DefaultCategoriser.cls) which is the default, provides the richest set of events. However it keeps track of the dimensions of the ListObject to achieve this -  if `Application.EventsEnabled = False` and a modification is made to the table then the categoriser gets out of sync and may miscategorise events.
 2. [NoCacheCategoriser](src/NoCacheCategoriser.cls) has a smaller subset of events which it raises, but is stateless so doesn't suffer from any cache invalidation if `Application.EventsEnabled` is toggled on/off. (Of course, no events will be triggered when `Application.EventsEnabled = False` by either categoriser)
 
