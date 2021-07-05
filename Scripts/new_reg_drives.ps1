New-PSDrive –Name HKCR –Root HKEY_CLASSES_ROOT –PSProvider Registry
New-PSDrive –Name HKU –Root HKEY_USERS –PSProvider Registry
New-PSDrive –Name HKCC –Root HKEY_CURRENT_CONFIG –PSProvider Registry
