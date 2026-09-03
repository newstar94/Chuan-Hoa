import win32com.client as win32

try:
    word = win32.GetActiveObject("Word.Application")
    print("Found active Word instance!")
    print("Documents count:", word.Documents.Count)
    print("Addins count:", word.AddIns.Count)
    for i in range(1, word.AddIns.Count + 1):
        try:
            addin = word.AddIns(i)
            print(f" - AddIn {i}: {addin.Name} (Installed: {addin.Installed}, Path: {addin.Path})")
        except Exception as e:
            print(f" - AddIn {i} error: {e}")
    
    print("COM Addins count:", word.COMAddIns.Count)
    for i in range(1, word.COMAddIns.Count + 1):
        try:
            ca = word.COMAddIns(i)
            print(f" - COMAddIn {i}: {ca.Description} (ProgID: {ca.ProgId}, Connected: {ca.Connect})")
        except Exception as e:
            print(f" - COMAddIn {i} error: {e}")
except Exception as ex:
    print("Error querying Word:", ex)
