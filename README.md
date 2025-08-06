# PPT Productivity Suite - Ribbon Update Instructions

This guide explains how to update your PowerPoint add-in to use the new ribbon UI, including the new "Paste Tools" tab and paste method buttons.

## Step 1: Replace the Ribbon.xml File

1. Open your project folder in Visual Studio or your preferred editor.
2. Locate the existing `Ribbon.xml` file in your project directory.
3. Replace its contents with the new ribbon XML provided below:

```
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Ribbon_Load">
	<ribbon>
		<tabs>
			<tab id="tabMain" label="Productivity Tools">
				<!-- ...existing groups and buttons... -->
			</tab>
			<tab id="tabPasteTools" label="Paste Tools">
				<group id="grpPasteMethods" label="Paste Methods">
					<button id="btnPastePlainText" label="Paste as Plain Text" onAction="OnPastePlainText" imageMso="PasteText" size="large"/>
					<button id="btnPasteDestinationTheme" label="Paste with Destination Theme" onAction="OnPasteDestinationTheme" imageMso="PasteDestinationTheme" size="large"/>
					<button id="btnPasteSourceFormatting" label="Paste with Source Formatting" onAction="OnPasteSourceFormatting" imageMso="PasteSourceFormatting" size="large"/>
					<button id="btnPasteImage" label="Paste Image" onAction="OnPasteImage" imageMso="PastePicture" size="large"/>
				</group>
			</tab>
		</tabs>
	</ribbon>
</customUI>
```

## Step 2: Ensure Resource Embedding (if required)

- If your add-in loads the ribbon from an embedded resource, make sure the updated `Ribbon.xml` is marked as an embedded resource in your project properties.
- In Visual Studio: Right-click `Ribbon.xml` > Properties > set "Build Action" to "Embedded Resource".

## Step 3: Implement Button Handlers

- The new buttons require handler methods in your ribbon controller (e.g., `RibbonController.cs`).
- Example handler stubs:

```csharp
public void OnPastePlainText(IRibbonControl control) { /* ... */ }
public void OnPasteDestinationTheme(IRibbonControl control) { /* ... */ }
public void OnPasteSourceFormatting(IRibbonControl control) { /* ... */ }
public void OnPasteImage(IRibbonControl control) { /* ... */ }
```

## Step 4: Build and Test

### How to Rebuild Your Add-in Project

1. **In Visual Studio:**
   - Open your solution.
   - Go to the menu bar and select **Build** > **Rebuild Solution**.
   - Wait for the build process to finish and check for any errors.

2. **In Visual Studio Code (with MSBuild or .NET CLI):**
   - Open the integrated terminal.
   - Run one of the following commands:
     ```
     dotnet build --no-incremental
     ```
     or
     ```
     msbuild /t:Rebuild
     ```
   - Ensure the build completes without errors.

3. Launch PowerPoint and verify the new "Paste Tools" tab and buttons appear.
4. Test each button to ensure it triggers the correct handler.

## Troubleshooting

- If the new tab does not appear, ensure the XML is valid and the file is embedded correctly.
- Check that your add-in is loading the correct resource.

---

For further customization or help with paste logic, see the source code or contact the project maintainer.
