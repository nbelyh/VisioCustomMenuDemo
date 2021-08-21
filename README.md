# VisioCustomMenuDemo

This sample project shows how one can add a custom **context menu** to a Visio VSTO add-in.

- You start off with "New Project" => "Visio VSTO Add-in".
- Then you add a Ribbon using right click on project, then Add => New Component => Ribbon (as XML)
- Then you can modify the Ribbon1.xml and Ribbon1.cs as in this project

The points that are demonstrated by this project:
- calling the add-in code when user click the menu item (specified in the "onAction" attribute)
- disabling/enabling/hiding menu items depending on selection
- adding image to a menu item

Running the thing (Microsoft Visio desktop must be installed):
- clone the repository
- open it in Visual Sutiodio
- Click "Run"

You should see a ribbon button added by the demo (Add-ins => Click Me)
If you click it, it creates a sample diagram (where you can test the context menus)

![image](https://user-images.githubusercontent.com/528366/130318567-bcb8fdeb-ddce-4315-9fb0-c00643d49d6d.png)

The demo menu item ("Rectangle") just paints the selected shape to red, and adds some text to it.
