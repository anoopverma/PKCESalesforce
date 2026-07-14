# Change Document - Salesforce OAuth2 PKCE Authenticator in VBA

This file tracks all modifications, enhancements, testing records, and releases of the Salesforce OAuth2 PKCE Authenticator library.

## Change History

### [2026-07-14] Initial Release & Architecture Setup
- Created repository design specification: [VBA_Salesforce_PKCE_Design.md](file:///c:/Users/703134581/projects/PKCE%20Connector%20%20Salesforce/docs/VBA_Salesforce_PKCE_Design.md)
- Created project tracking structure.
- Initialized local Git repository, configured remote pointing to `https://github.com/anoopverma/PKCESalesforce`.

### [2026-07-14] Added Cryptography & Parser Modules
- Implemented [CryptoUtils.bas](file:///c:/Users/703134581/projects/PKCE%20Connector%20%20Salesforce/src/CryptoUtils.bas): Added secure random bytes generator via Windows BCrypt API, UTF-8 encoder, SHA-256 hash generator, and Base64/Base64URL encoders.
- Implemented [JSONParser.bas](file:///c:/Users/703134581/projects/PKCE%20Connector%20%20Salesforce/src/JSONParser.bas): Implemented a native recursive-descent JSON parser mapping objects to Scripting.Dictionary and arrays to Collection.

### [2026-07-14] Added OAuth2 PKCE Flow & Login Form
- Implemented [SalesforceOAuth2.bas](file:///c:/Users/703134581/projects/PKCE%20Connector%20%20Salesforce/src/SalesforceOAuth2.bas): Handles URL encoding, assembling authorize URL, capturing and parsing auth codes (supporting manual redirect fallback), and token requests (both code exchange and token refresh).
- Implemented [frmOAuthLogin.frm](file:///c:/Users/703134581/projects/PKCE%20Connector%20%20Salesforce/src/frmOAuthLogin.frm): Created code-behind for the optional login UserForm utilizing a dynamically instantiated WebBrowser control to capture auth codes interactively.

### [2026-07-14] Added API Integration & Tests Modules
- Implemented [SalesforceAPI.bas](file:///c:/Users/703134581/projects/PKCE%20Connector%20%20Salesforce/src/SalesforceAPI.bas): Added core REST API executor supporting GET/POST/PATCH methods, custom authorization headers, and SOQL query helper function.
- Implemented [OAuthDemo.bas](file:///c:/Users/703134581/projects/PKCE%20Connector%20%20Salesforce/src/OAuthDemo.bas): Created example Excel macros for initiating PKCE flow, extracting tokens, and running a test SOQL query.
- Implemented [VBA_Tests.bas](file:///c:/Users/703134581/projects/PKCE%20Connector%20%20Salesforce/src/VBA_Tests.bas): Built custom test suite verifying correct outputs for SHA-256 (via test vectors), Base64URL, code verifier generation, and recursive JSON parsing.
