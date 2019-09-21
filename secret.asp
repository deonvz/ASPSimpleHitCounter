<html>
<head>
<title></title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="_themes/ccs/ccs1000.css"><meta name="Microsoft Theme" content="ccs 1000, default">
</head>

<table>
<tr>
<td>
<!--#include file="banner.html" -->
</td>
</tr>
<tr>
<td>

<body>
<!---
	Created by: Deon van Zyl
	Date: 2000
	Simple Counter - Free to use
-->
<p align="left">
<%

' Declare variables
Dim ObjCounterFile, ReadCounterFile, WriteCounterFile
Dim CounterFile
Dim CounterHits

Set ObjCounterFile = Server.CreateObject("Scripting.FileSystemObject")

	CounterFile = Server.MapPath ("candycounter.txt")

	Set ReadCounterFile= ObjCounterFile.OpenTextFile (CounterFile, 1, True)

		If Not ReadCounterFile.AtEndOfStream Then
			CounterHits = Trim(ReadCounterFile.ReadLine)
			If CounterHits = "" Then CounterHits = 0
		Else
			CounterHits = 0
		End If

	ReadCounterFile.Close
	Set ReadCounterFile = Nothing

	CounterHits = CounterHits + 1

	Set WriteCounterFile= ObjCounterFile.CreateTextFile (CounterFile, True)
		WriteCounterFile.WriteLine(CounterHits)
	WriteCounterFile.Close
	Set WriteCounterFile = Nothing

Set ObjCounterFile = Nothing
%>

<br>
This page has been accessed : <b><Font face="arial" color="Yellow" size="2"><% =CounterHits %></font></b> since Feb 2004
</td>
</tr>
<tr>
<td>
&nbsp;
</td>
</tr>
<tr>
<td>
Introduction to Candy can be found <a href="Candy 00 - Introducing Candy rev3.pdf">here</a>.
</td>
</tr>
</table>
</body>
</html>