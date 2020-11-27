<%--
  Created by IntelliJ IDEA.
  User: zhangyuan
  Date: 2020/10/19
  Time: 下午6:19
  To change this template use File | Settings | File Templates.
--%>
<!DOCTYPE html>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
  <link href='//fonts.googleapis.com/css?family=Marmelad' rel='stylesheet' type='text/css'>
  <title>Hello App Engine Standard Java 8</title>
</head>
<body>
<h1>Hello App Engine -- Java 8!</h1>

<p>This is appengine for Java 8</p>
<table>
  <tr>
    <td colspan="2" style="font-weight:bold;">Available Servlets:</td>
  </tr>
  <tr>
    <td><a href='/_ah/admin'>Datastore viewer</a></td>
  </tr>
  <tr>
    <td><input id="kind" placeholder="请输入需导出kind的名称"/></td><td><button onclick="onClick()">export</button></td>
  </tr>
</table>
</body>
<script>
  function onClick() {
    var kind = document.getElementById('kind').value;
    if (!kind) {
      alert("kind名称不能为空");
      return;
    }
    window.location.href = "/downloadKind?kind=" + kind;
  }
</script>
</html>
