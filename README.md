# Plugin Bankimport

This is a plugin for [www.openpetra.org](http://www.openpetra.org).

## Functionality

You can import bank statements into OpenPetra, and match the donors and Motivation Detail to a gift, and allows to create Gift Batches.
It will also possible to do the same for GL Batches.

When you import the bank statements in the next month, the plugin will remember the matches of recurring donations, and match the new donations to the same donor and project.

This plugin uses the tables [a_ep_statement] (https://ci.openpetra.org/job/OpenPetraDBDoc/doclinks/1/index.html?table=a_ep_statement&group=account) and [a_ep_transaction] (https://ci.openpetra.org/job/OpenPetraDBDoc/doclinks/1/index.html?table=a_ep_transaction&group=account) for storing the bank statements,
and the table [a_ep_match] (https://ci.openpetra.org/job/OpenPetraDBDoc/doclinks/1/index.html?table=a_ep_match&group=account) for storing the matches.

## Dependencies

This plugin requires one of the following plugins:

* https://github.com/SolidCharity/OpenPetraPlugin_BankimportCSV
* https://github.com/SolidCharity/OpenPetraPlugin_BankimportMT940

## Installation

Please copy this directory to your OpenPetra working directory, to csharp\ICT\Petra\Plugins, or include it like this, if you are using git anyway:

    git submodule add https://github.com/SolidCharity/OpenPetraPlugin_Bankimport csharp/ICT/Petra/Plugins/Bankimport

also apply the patch GiftBatchExport.patch!
    
and then run

    nant generateSolution

Please check the config directory for changes to your config files.

## License

This plugin is licensed under the GPL v3.