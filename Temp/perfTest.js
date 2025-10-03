//[group=Temp]
Repository.EnsureOutputVisible("Script")
var curDG = Repository.GetCurrentDiagram()
if (curDG != null){
	var start = new Date().getTime()
	Repository.SaveDiagram(curDG.DiagramID)
	var stop = new Date().getTime()
	Repository.WriteOutput("Script", "Time: " + (stop - start) + " ms", 0)
}