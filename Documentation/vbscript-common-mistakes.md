# VBScript Common Mistakes

## Incorrect use of parentheses for Sub

See [Argument in Parentheses](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/ee478101(v=vs.84)#argument-in-parentheses) for more details.

You will get confused because sometimes you can call a procedure with parentheses, that is when there are zero or one arguments in the argument list.

In the case of one argument what is happening is that the argument is being passed `ByVal` instead of `ByRef`.

Summary:

A function call uses parentheses.

A procedure call does not use parentheses.

## AND and OR do not short-circuit

[And](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/e8zy95hw(v=vs.84)) and [OR](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/w6a4aywe(v=vs.84)) logical operators do not short-circuit.

More modern language have logical operators that short-circuit, which means that as soon as one expression determines the answer the rest of the tests are skipped.

**WRONG**
```
    ' This is WRONG and will fail at runtime when tvArchimateStyleColor is nothing
    ' as tvArchimateStyleColor.Value will be called
    set tvArchimateStyleColor = taggedValues.GetByName("ArchiMate::Style::Color")
    if not tvArchimateStyleColor is nothing and tvArchimateStyleColor.Value = "ignore" then
```

**CORRECT**
```
    ' This is the CORRECT handling when a variable could be nothing
    set tvArchimateStyleColor = taggedValues.GetByName("ArchiMate::Style::Color")
    if not tvArchimateStyleColor is nothing then
        if tvArchimateStyleColor.Value = "ignore" then
```
