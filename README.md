# DFIR-PowerShell
Collection of PowerShell DFIR Scripts.
## Get-RecentFileHashes
This script pulls recently opened file history from Internet Exploer cache, and calculates the hash for the file supports SHA1, SHA256, SHA384, SHA512, MACTripleDES, MD5, and RIPEMD160.  Two CSV's are created; a raw IE file history and a cleaned up file with the hashes.
