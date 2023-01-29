<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    <link rel="stylesheet" href="https://tistory1.daumcdn.net/tistory/5499297/skin/style.css">
    <link rel="stylesheet" href="https://tistory1.daumcdn.net/tistory/5925927/skin/style.css">
    <link rel="stylesheet" href="./asciiartanieman.css">
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.6.1/jquery.min.js"></script>

</head>

<body>
    <div>
        <header>
            <h2 onclick="this.style.background='var(--orange)'">게시물 목록</h2>
        </header>
    </div>
    <%
        dim fs,fo,x
        set fs=Server.CreateObject("Scripting.FileSystemObject")
        set fo=fs.GetFolder("C:\Users\L340-15API\24\aacss\posts")

        for each x in fo.files
            'Print the name of all files in the test folder
            Response.write("<a href=http://chonggi-tokhu.r-e.kr/posts/" & x.Name & " class=colourbluen>" & x.Name & "</a>" & "<p><br></p>")
        next

        set fo=nothing
        set fs=nothing
    %>
</body>

</html>