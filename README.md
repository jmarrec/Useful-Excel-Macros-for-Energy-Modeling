# Goal

* Some shortcuts
* A macro to export an Excel table to an IDF format for import into IDF Editor or to paste in the idf file directly
* A macro to export an excel table to a Ruby array of hash: useful when working with ruby and the OpenStudio bindings especially
* A macro to export an excel table to JSON
* A macro to generate a ruby script to create curves using the OpenStudio ruby bindings, using IDF objects as input
* Psychrometrics function to calculate the humidity ratio from RH(%) to Dry Bulb Temperature in SI units, useful to model Coil:Cooling:Water

# Instructions

I. Install the Add-in.
----------------------

1. Put Add-in (Useful_Macros_for_BEM.xlam) in your Appdata folder.
To do so, open the Run window by pressing [Windows + R] and type in: %appdata%\Microsoft\AddIns

2. Open Excel, go to File > Options > Add-ins > Click on Useful_Macros_for_BEM and then "Go".

3. In the Dialog Box choose Useful_Macros_for_BEM.

(4. If you open VBE (Visual Basic Editor, Alt+F11), you should see it. You can close VBE)

Optional: add buttons in the ribbon for easy access

5. Right click on the ribbon and click "Customize the ribbon".
You can create on the right side a New Tab ("My Macros" For example), a new group ("BEM" For example).
On the right side, at the top, "Choose commands from" and select Macro. Find "Export_To_IDF" and put it in the group you just created.
Click on "Rename". "Export_To_IDF" can be called "Export to IDF" for example.
(You can also choose a pretty icon like an up arrow for example...)

6. Repeat 5. for "Export_To_JSON_Array_Of_Hash", etc


## Excel to IDF

### Usage 

For exporting from IDF to IDF Editor or idf file, I wrote a macro that will read in an excel table and export that to an IDF compatible object so that you can paste it in the IDF editor or in your .idf text file itself.

This is related to my answer on Unmet Hours on the post [EnergyPlus IDF editor copy paste](https://unmethours.com/question/17809/energyplus-idf-editor-copy-paste). I wrote the macro more than 2 years ago for the same purpose, but only recently realized it could be useful to others.

I've personally added this macro to an Excel Add-In and added a button on a separate tab for easy access:

![Excel to IDF Button](/doc/Excel_to_IDF_button.png)

To use the macro, place your cursor in any of the cells of the table and launch it. It will try to find the EPlus class name in the cell that's 2 rows above the top-left corner of the table to prepopulate a dialog box asking you for the class name: either it's good and click OK, or it isn't or it didn't find it, type the name of the class.

It'll then generate proper IDF objects that you can paste in IDF editor or paste directly in the .idf file. It'll ask you whether you want to save it as a file (and subsequently ask for a file name) or copy it to clipboard.

### Example

Here's an example:

![Example Excel to IDF](/doc/Excel_to_IDF_dialog_box.png)

And Here's the output it produced for the example table:

    Eplus:ObjectClass,
        Object 1,
        (1,1),
        (1,2),
        (1,3);

    Eplus:ObjectClass,
        Object 2,
        (2,1),
        (2,2),
        (2,3);

    Eplus:ObjectClass,
        Object 3,
        (3,1),
        (3,2),
        (3,3);


## Export table to JSON array of Hash
--------------------------------------

Pretty much the same idea as above, you can place yourself anywhere in the table and run the macro.

**Important: Do not include any white space in your table headers: I'm using the `:key` notation in the hash

If you are using a number it'll be formatted as a number in the hash. Anything else will be a string.

### Example

Here's an example table

![Json example table](/doc/excel_to_json_example.png)

And here's the output:

    myhash = [{:zone => 'Zone 1',:supplyAirFlowrate => 18.17,:cooling_cop => 4.88,:fan_pressure_rise => 714.82,:total_fan_eff => 0.605,},
    {:zone => 'Zone 2',:supplyAirFlowrate => 13.988,:cooling_cop => 5.11,:fan_pressure_rise => 826.51,:total_fan_eff => 0.514,},
    {:zone => 'Zone 3',:supplyAirFlowrate => 2.401,:cooling_cop => 5.11,:fan_pressure_rise => 826.51,:total_fan_eff => 0.514,},
    {:zone => 'Zone 4',:supplyAirFlowrate => 5.805,:cooling_cop => 5.24,:fan_pressure_rise => 756.96,:total_fan_eff => 0.644,},
    {:zone => 'Zone 5',:supplyAirFlowrate => 6.072,:cooling_cop => 5.75,:fan_pressure_rise => 896.4,:total_fan_eff => 0.474,},
    {:zone => 'Zone 6',:supplyAirFlowrate => 6.072,:cooling_cop => 5.75,:fan_pressure_rise => 896.4,:total_fan_eff => 0.474,},
    ]

## OpenStudio Ruby Curve Creator.

From a table taken from the IDF editor, it will generate Ruby code to generate the curves.

I added a spreadsheet that includes instructions and that allows you to generate all types of curves in one go. But I also added it in the add-in, because sometimes you don't want to spin a workbook to do a task.

Here's an example of a table with two Quadratic curves.

![OpenStudio Ruby Curve Creator](/doc/OpenStudioRubyCurveCreator.png)

And here's the code it generates:

```
curve = OpenStudio::Model::CurveQuadratic.new(model)
curve.setName('CurveQuad1')
curve.setCoefficient1Constant(0)
curve.setCoefficient2x(1)
curve.setCoefficient3xPOW2(0)
curve.setMinimumValueofx(0.03)
curve.setMaximumValueofx(1)

curve = OpenStudio::Model::CurveQuadratic.new(model)
curve.setName('CurveQuad2')
curve.setCoefficient1Constant(0.7516)
curve.setCoefficient2x(0.00414)
curve.setCoefficient3xPOW2(0)
curve.setMinimumValueofx(29.44)
curve.setMaximumValueofx(85)
curve.setInputUnitTypeforX('Temperature')
curve.setOutputUnitType('Dimensionless')
```


## Shortcuts

In "ThisWorkbook" you have shortcuts that I use on a daily basis. Feel free to adapt or delete completely.

* `CTRL` + `SHIFT` + `F1` = Change Reference style to R1C1
    
* `ALT` + `LEFT` = Decrease number of Decimal places by 1
    
* `ALT` + `RIGHT` = Increase number of Decimal places by 1
    
* `CTRL` + `ALT` + `C` = Center cell horizontally and vertically

* `CTRL` + `SHIFT` + `P` = Paste Values Only



### Contact and Contribution

**Happy modeling and don't hesitate to reach out to me for any bugs or comments, using the "issues" tab preferably.**

I'll also welcome pull requests.