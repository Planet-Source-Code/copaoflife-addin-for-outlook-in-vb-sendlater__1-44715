--------------------------------------------------
Creating the addin - sendLater!
----------------------------------------------------------------------------------------------------
Create a new 'Addin' project in Visual Basic. 
Visual basic wires all code that is required assuming that you are developing the addin for VB itself.. 

So u have to remove some code and add some to make it a pure outlook addin.

U can happily remove the form that vb has created.. if ur addin doesnt need one.. in this example though, the form is modified to fit the usage.

open the designer, and u will find two tabs.
U can enter the display name for ur addin, with a description. 

U can select the target application for which ur a developing the addin. for our example, it would be, Microsoft outlook, and version is 9.0

'Initial Load Behaviour' is the property which sets how the addin is loaded on to the target application. 
Leave it at 'startup' - This means that outlook would load the addin by default everytime outlook 
starts.


In this application we wont modify anything in the advanced tab. 

Now switch to 'code view' in vb, to see the code of the designer.

U can remove referrences of 'Vb 6.0 Extensibility', and 'Microsoft office 10.0 object library' from project referrences.
add 'microsoft outlook 9.0 object library' to the referrences.

remove all code related to the referrences u just removed from the project. now create an instance of the outlook object, and wire ur code.


To debug or walk-through the code to understand it, follow these steps.

1 close outlook if its open.
2 keep a break point in the onconnection event in the 'connect' object/designers code.
3 then 'Run' the project. since an addin project is similar to a dll project, it will be in run mode. 
4 open outlook.. as soon as u open, the vb code that was running reaches to the break point and waits.
outlook wont complete loading until u run the vb code again beyond the breakpoint.


once we run the addin code, which is nothing but a dll code, it gets registered as an outlook dll in the registry. vb takes care of this wiring. and when outlook starts, and since we mentioned load behaviour as 'startup' in the designer, outlook will detect the addin and try to connect to it... and thus we reached the breakpoint.

The rest of it is self-explanatory...i believe.. when u explore the code.

This application/addin was developed and tested on windows2000, with outlook 2000 and Visual Basic 6.0.

This application uses microsoft windows common control 2.6.0 (which comes with visual studio sp4), for the datepicker control.

The code also uses the following referrences.
Microsoft Add-In Designer, 
Microsoft Outlook 9.0 Object Library.

----------------------------------------------------------------------------------------------------
To just use the application on your outlook.., register the dll provided with the code(.zip file) on your machine.. and open outlook .... >> options >> other >> advanced >> comaddins >> add new 'addin', point to the dll.
--------------------------------------------------
Then open a new mail window and try sending a mail, to see the dll in action.