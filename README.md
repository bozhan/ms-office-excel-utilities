# ms-office-excel-utilities
Utitlity classes for MS Excel version 2003-2019
This repository provides a utility class bundle written in VBA that will help you automate complex tasks such as importing, exporting and syncing with other sheets. Doing file management with ease. Looking up and resolving contacts in outlook and active directory and much more.
You will find an excel sheet that includes all my utility classes and puts them to practical use.
In addition the utility and all other helper classes are exported in their own files in the Excel Utilities.xlsm_Files folder.
These are classes that build on top of the functionality and abstractions that come with VBA.

The main interface class Utility exposes all other helper and utility classes. This is done purely out of convinience for the developer. All utility classes are initiated in the interface class constructor and thus available when the document is opened.
This allows the develper to refer them in the debug console and in any other class just by navigating through Util.<class> .
