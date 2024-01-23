# Copy here the e-mail attachments input from Power Automate action
$JSONOutput = '[
  {
    "Name": "K461-KT-OMV-T-00283.pdf",
    "ContentBytes": "JVBERi0xLjUNCiW1tbW1DQoxIDAgb2JqDQo8PC9UeXBlL0NhdGFsb2cvUGFnZXMgMiAwIFIvTGFuZyhlbi1VUykgL1N0cnVjdFRyZWVSb290IDE4IDAgUi9NYXJrSW5mbzw8L01hcmtlZCB0cnVlPj4+Pg0KZW5kb2JqDQoyIDAgb2JqDQo8PC9UeXBlL1BhZ2VzL0NvdW50IDIvS2lkc1sgMyAwIFIgMTYgMCBSXSA+Pg0KZW5kb2JqDQozIDAgb2JqDQo8PC9UeXBlL1BhZ2UvUGFyZW50IDIgMCBSL1Jlc291cmNlczw8L0ZvbnQ8PC9GMSA1IDAgUi9GMiA5IDAgUi9GMyAxNCAwIFI+Pi9FeHRHU3RhdGU8PC9HUzcgNyAwI..."
  }
]'

# Convert JSON to Powershell object
$pdf_object = ConvertFrom-Json $JSONOutput

# Full path destination to the file to be created
$NewFilePath = "$pwd\$($pdf_object.Name)"

# Convert from Base64
$pdf_bytes = [System.Convert]::FromBase64String($pdf_object.ContentBytes)

#Create file
[IO.File]::WriteAllBytes("$NewFilePath", $pdf_bytes)