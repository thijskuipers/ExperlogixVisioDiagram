# Experlogix to Visio diagram generator

This tool allows you to create a Visio dependency diagram from a particular Series/Model in Experlogix. It draws:

* Lists
* Lookup tables
* Categories (with attributes)
* Formulas
* Rules

The tool does **not** resolve dependencies to Options (e.g. salesnumbers in the *premise* or *conclusion* of a Rule) or dynamic properties in formulas (e.g. `[.Quantity]`).

This tool is inspired by the "metadatadiagram" tool of the Microsoft Dynamics CRM 2011 SDK.

## How to use

Locate the Experlogix database (it's a \*.mdb file) Make sure the Experlogix database is not in use. It is in use when there is a \*.ldb file located next to it. Make sure you've closed the Experlogix configuration application.

### Configuration

You can either:

#### Move and rename the database

Copy the database file to the application folder and rename it to "Experlogix.mdb".

#### Edit the configuration file

Edit the Broes.Experlogix.VisioDiagram.exe.config file to point to the correct location of the
Experlogix MDB file:

    <add name="Broes.Experlogix.DAL.Properties.Settings.ExperlogixConnectionString"
        connectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=&quot;Experlogix.mdb&quot;"
        providerName="System.Data.OleDb" />

It's the bit after the `Data Source=` between the two `&quot;`'s. Make sure you keep the two `&quot;`'s!

### Run the application

To run the application, open the Broes.Experlogix.VisioDiagram.exe file.

#### Choose Series

If the application is configured correct, it will ask you to choose a Series:

    [1] Series 1
    [2] Series 2
    [3] Series 3
    Choose the series [1 - 3]: 

Enter a number and press Enter.

#### Choose Model

It will ask you to choose a Model:

    [1] Model 1
    [2] Model 2
    Choose the model [1 - 2]: 

Enter a number and press Enter

#### Output

The application will now draw the dependency diagram and will give feedback on its progress:

    Drawing 38 categories...
    Drawing 29 rules...
    Drawing 12 lists...
    Drawing 3 lookup tables...
    Drawing 67 formulas...
    Drawing 341 relations...
    * Unable to draw from FOR_RetrieveClientSizeDiscoun to CAT_MPC
    * Unable to draw from FOR_RetrieveAreaDiscount to CAT_MPC
    
    Laying out the page...
    Resizing to fit to contents...
    
    Press enter to exit...

Error messages are a convenient way to find errors in the dependency graph. This tool does a (limited) static analysis of the dependencies within a model.

If the tool ran successfully, you will see a Visio document with the generated dependency diagram:

![Experlogix to Visio diagram result](http://www.broes.nl/wp-content/uploads/2012/10/ExperlogixVisioDiagram-result.png)

## Prerequisites

* [Experlogix](http://experlogix.com/) database (*.mdb) (tested on Experlogix database version 6.5.K, see table "Version" in database)
* [.NET Framework 4.5](http://www.microsoft.com/en-us/download/details.aspx?id=30653)
* Microsoft Visio (2007 and 2010 tested)
* In case you did not install Visual Studio and/or its Microsoft Office Developer Tools, you'll need to install the Microsoft Office Primary Interop Assemblies. Depending on your Office version, pick one of the following:
  * [Microsoft Office 2010: Primary Interop Assemblies Redistributable](http://www.microsoft.com/en-us/download/details.aspx?id=3508)
  * [2007 Microsoft Office System Update: Redistributable Primary Interop Assemblies](http://www.microsoft.com/en-us/download/details.aspx?id=18346)
  * [Office 2003 Update: Redistributable Primary Interop Assemblies](http://www.microsoft.com/en-us/download/details.aspx?id=20923)
