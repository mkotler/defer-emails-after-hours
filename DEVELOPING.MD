# Development Instructions

During development follow these steps to get the add-in to work.  

1. Build the add-in:  `npm run build`
2. Run `npm run validate` to make sure the manifest.xml is valid
3. Start the dev server: `npm run dev-server`
4. Navigate to [https://localhost:3000/manifest.xml](https://localhost:3000/manifest.xml)
5. Right click in the browser and select "Save As"
6. Save the file with a ".xml" extension
7. Navigate to [Add-ins for Outlook](https://aka.ms/olksideload)
8. Select My add-ins
9. In the "Custom Addins" section (at the bottom of the page) select "Add a custom add-in" and choose "Add from file"
10. Select the XML file saved in step 6 and click "Install"
11. Navigate to Outlook and refresh the browser