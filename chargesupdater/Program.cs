using System;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

Console.WriteLine("To start, remove all charges from config.lua. It should start at line 40. Then replace the newcharges.csv. IT MUST MATCH THAT FILE NAME EXACTLY ELSE IT WILL NOT WORK. Hit any key when that's done");
Console.ReadKey();
string strExeFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
string strWorkPath = System.IO.Path.GetDirectoryName(strExeFilePath);

string charges = strWorkPath + @"\newcharges.csv";
string config = strWorkPath + @"\config.lua";
var lines = File.ReadAllLines(config).ToList();
//lines.Insert(40, "@  ['TestCharge'] = {label = 'whatever', jail = 48, fine = 19, color = '#ff2e2e'},");
// lines.Insert(skip, $"  ['{id}'] = {{label = {label}, jail = {time}, fine = {fine}, color = '#ff2e2e'}},");
var chargesLines = File.ReadAllLines(charges);
var skip = 40;
var iColor = "#93c47d";
var mColor = "#f6b26b";
var fColor = "#e06666";
var vfColor = "#990000";
var sfColor = "#990000";
var color = "#ffffff";
foreach (var line in chargesLines) {
    var vals = line.Split(',');
    var id = vals[0];
    var label = vals[1];
    switch (vals[2]) {
        case "I":
            color = iColor;
            break;
        case "M":
            color = mColor;
            break;
        case "F":
            color = fColor;
            break;
        case "VF":
            color = vfColor;
            break;
        case "SF":
            color= sfColor;
            break;
    }
    var jail = vals[4];
    var fine = vals[3];
    lines.Insert(skip, $"  ['{id}'] = {{label = '{label}', jail = {jail}, fine = {fine}, color = '{color}'}},");
    skip++;
    
}

File.WriteAllLines(config, lines);
    
