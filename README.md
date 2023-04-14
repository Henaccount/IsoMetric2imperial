# IsoMetric2imperial
this is example code, use at own risk. It is showing the basic tasks to replace text in certain dimensions/texts/table cells for a batch of isometric DWGs. 
It is not expected to completely cover all objects containing dimensional text yet, but of course you can extend the code to cover more.
It creates copies of Plant 3D isometrics and replace certain metric values with imperial values.
Load the dll with "netload". Run the script with the command: "IsoMetric2imperial". A file dialog will open and you need to select one DWG of a folder, but the script will execute over all DWG in this folder (not in the subfolders). The script runs silently with no output other than "script execution ended". The modified copies of the DWGs will be stored in a subfolder called "results". There will be also a logfile showing potential errors.

