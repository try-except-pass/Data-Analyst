# data-analyst
VBA scripting for daily tasks as a Commissions Specialist
Requirements:

Personal worksheet
Macros enable in Trust Center
Having reference "Microsoft Scripting Runtime" enabled
VBA script with user interface which allows the user to:

Create a Pivot generating macro Macro adds a worksheet to the Personal file with the title of the macro, with 4 columns: Filters, Rows, Columns, Data. Using these values, the macro will create a pivot in a new worksheet with the fields choosen by the user.

Edit existing macros User can select which Macro to edit in the ComboBox. By using two ListBoxes for each each column, user can choose the fields they want to add from the fields available in the data export where the macro is ran. User can also define the order of the fields, all of which achieved with command buttons (Add, Up, Down, Delete). Every input is immediately saved in the personal file.

Delete existing macros

Alongside business changes, new and different data exports are used, causing macros to become obselete or outdated, as column headers change, are no longer needed, or a better configuration is found. The idea was to create an self managed and future proof macro, where the user with no knowledge of scripting, will be able to create, adapt, and manage their macros for their current reporting needs.

Selection view:
![image](https://github.com/try-except-pass/data-analyst/assets/73493873/1a1e1ee9-1587-4382-9363-76f1d0d39e9c)


Editing view:
![image](https://github.com/try-except-pass/data-analyst/assets/73493873/424df1d9-2511-417d-94a9-2612b9ee7332)
