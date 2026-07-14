# Integration & Usage Guide: Salesforce PKCE Authenticator in VBA

This guide details how to import, reference, and use the Salesforce PKCE Authenticator library in your Excel Workbooks.

---

## Method 1: Direct Import (Recommended)

This is the simplest way to add the authenticator to an existing Excel workbook.

### Step 1: Open the VBA Editor
1. Open your Excel workbook.
2. Press **`Alt + F11`** to open the Visual Basic for Applications (VBA) Editor.

### Step 2: Enable Required References
This library requires the **Microsoft Scripting Runtime** (for dictionary support) and **Microsoft XML, v6.0** (for HTTP requests).
1. In the VBA Editor menu, go to **Tools > References...**.
2. Scroll down and check the box for:
   - **`Microsoft Scripting Runtime`**
   - **`Microsoft XML, v6.0`**
3. Click **OK**.

### Step 3: Import the Source Files
You need to import the `.bas` modules you generated:
1. In the **Project Explorer** (usually on the left side; press `Ctrl + R` if hidden), right-click your workbook name.
2. Select **Import File...**.
3. Navigate to the folder where you cloned the repository (`src/` directory) and import:
   - `CryptoUtils.bas`
   - `JSONParser.bas`
   - `SalesforceOAuth2.bas`
   - `SalesforceAPI.bas`
   - `OAuthDemo.bas`
   - `VBA_Tests.bas`
   - *(Optional)* `frmOAuthLogin.frm` (This will automatically import its companion `.frx` layout file if available).

---

## Method 2: Creating an Excel Add-In (`.xlam`)

If you want to use these functions across *multiple* workbooks without importing the files into each one, you can package the modules into an **Excel Add-In**.

### Step 1: Prepare the Add-In Workbook
1. Create a fresh, blank Excel workbook.
2. Import all the `.bas` files and enable the references using **Method 1** above.
3. Save the workbook as an Excel Macro-Enabled Workbook (`.xlsm`) first, named `SalesforcePKCE.xlsm`.

### Step 2: Save as an Add-In
1. Go to **File > Save As**.
2. In the "Save as type" dropdown, select **Excel Add-In (*.xlam)**.
3. Excel will automatically change the save location to your local AddIns folder (usually `AppData\Roaming\Microsoft\AddIns`).
4. Click **Save** and close Excel.

### Step 3: Enable the Add-In in Excel
1. Open any Excel workbook where you want to use the library.
2. Go to **Developer > Excel Add-ins** (or **File > Options > Add-ins > Manage: Excel Add-ins > Go**).
3. Check the box for **`SalesforcePKCE`** and click **OK**.

### Step 4: Reference the Add-In in VBA
To write code that calls the Add-In directly:
1. Open the VBA Editor (`Alt + F11`).
2. Go to **Tools > References**.
3. Check the box for **`SalesforcePKCE`** (it will be listed as a reference).
4. Now, in your macros, you can call any public function directly, for example:
   ```vba
   Dim verifier As String
   Dim challenge As String
   verifier = CryptoUtils.GenerateCodeVerifier()
   challenge = CryptoUtils.GenerateCodeChallenge(verifier)
   ```

---

## Running the Verification Tests
To verify that everything is working perfectly on your machine:
1. In the VBA Editor, open the Immediate Window (**`Ctrl + G`**).
2. Click inside any module, type the following command in the Immediate Window, and press Enter:
   ```vba
   VBA_Tests.RunAllTests
   ```
3. A success message box should appear confirming that SHA256, Base64URL, and the JSON Parser are running correctly on your Excel environment.
