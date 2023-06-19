<!-- #include file="uploadhelper.asp" -->
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title></title>
  </head>
  <body>
    <form action="<%=Request.ServerVariables("SCRIPT_NAME")%>?cmd=upload" method="POST" enctype="multipart/form-data">
      File: <input type="file" name="files[]" multiple>
      <input type="submit" value="Upload">
    </form> 
<%
If Request("cmd")="upload" Then
    Dim objUpload : Set objUpload = New UploadHelper
    If objUpload.GetError <> "" Then
        Response.Write("Warning: "&objUpload.GetError)
    Else  
        Response.Write("found "&objUpload.FileCount&" files...<br />")
        Dim x : For x = 0 To objUpload.FileCount - 1
            Response.Write("file name: "&objUpload.File(x).FileName&"<br />")
            Response.Write("file type: "&objUpload.File(x).ContentType&"<br />")
            Response.Write("file size: "&objUpload.File(x).Size&"<br />")
            Response.Write("Saved at: D:\TestASPProject\Photos<br />")
            'If want to convert the virtual path to physical path then use MapPath
            'Call objUpload.File(x).SaveToDisk(Server.MapPath("/public"), "")
            Call objUpload.File(x).SaveToDisk("D:\TestASPProject\Photos", "")
            Response.Write("file saved successfully!")
            Response.Write("<hr />")
        Next            
    End If
End If
%>
  </body>
</html>