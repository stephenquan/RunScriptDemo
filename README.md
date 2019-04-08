# RunScriptDemo

I've put together a "Hello World" IActiveScript C++ ATL console application that:

- Define CSimpleScriptSite class
  - Implement IActiveScriptSite interface (mandatory)
  - Implement IActiveScriptSiteWindow interface (optional)
  - Minimum implementation with most functions implemented with a dummy stub
  - Has no error handling. Consult MSDN IActiveScriptError.
- Use CoCreateInstance a new IActiveSite object
  - Create instances of both VBScript and JScript
  - Link the IActiveSite to IActiveScriptSite using IActiveSite::SetScriptSite
  - Call QueryInterface to get an IActiveScriptParse interface
  - Use IActiveScriptParse to execute VBScript or JScript code
- The sample:
  - Evaluates an expression in JScript
  - Evaluates an expression in VBScript
  - Runs a command in VBScript

References:
 - https://stackoverflow.com/a/9150038/881441
 - https://stackoverflow.com/questions/7491868/how-to-load-call-a-vbscript-function-from-within-c/9150038#9150038
