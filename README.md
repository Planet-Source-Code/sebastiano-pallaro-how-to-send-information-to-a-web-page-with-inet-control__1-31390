<div align="center">

## How to send information to a web page with Inet control


</div>

### Description

The purpose of this article is explain how to send informations to a web page with an HTML-Form and Inet control. Really easy.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sebastiano Pallaro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sebastiano-pallaro.md)
**Level**          |Intermediate
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0, ASP \(Active Server Pages\) 
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sebastiano-pallaro-how-to-send-information-to-a-web-page-with-inet-control__1-31390/archive/master.zip)





### Source Code

<p class="MsoNormal"><font face="Tahoma" size="2"><b>How to send information to
a web page with Inet control<br>
</b><br>
Sending information to a web page with VB is really simple. You must only use
the Inet control and you application will be able to communicate with a web
site.<br>
On this small tutorial I will suppose that we need an is-update-aviable-check
routine.<br>
The first step is to build the web page that must receive the information. The
page must know the version number of the program and, in addition, the
registration-key of the program.<br>
Here is the ASP page that check the version:</font></p>
<p class="MsoNormal"><font face="Courier New" size="2"><font color="#800000"><span style="background-color: #FFFF00">&lt;%</span></font><span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp; </span><font color="#008000">‘ Suppose that “B” is the
new version of our application.</font><span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp; </span><font color="#0000FF">if </font>request.<b>form</b>(“<b>Version</b>”)
&lt;&gt; ”<b>B</b>” <font color="#0000FF">then</font><span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>response.write “<b>to_app</b>”
<font color="#008000">‘ Hey, you must update!</font><span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp; </span><font color="#0000FF">else</font><span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>response.write “<b>ok</b>”
<font color="#008000">‘ All ok.</font><span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp; </span><font color="#0000FF">end if</font><span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp; </span><font color="#008000">‘ With registration key you
can do all you want, for example put it into a database to track users<span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp; </span>‘ updates… he he he good for privacy :-]<span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp; </span>‘ …</font><br>
<font color="#800000"><span style="background-color: #FFFF00">%&gt;</span></font></font><br>
<br>
<font face="Tahoma" size="2">Now we only need to add to our program the
check-routine.<br>
On a form put an Inet control (Inet1)<span style="mso-spacerun: yes">&nbsp; </span>and
add this code, for example, on the click event of a button:<span style="mso-spacerun: yes">&nbsp;&nbsp;</span></font></p>
<p class="MsoNormal"><font face="Courier New" size="2"><span style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;
</span><font color="#008000">‘ …</font><span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp; </span><font color="#0000FF">Dim</font> strUrl <font color="#0000FF">As
String</font><span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp; </span><font color="#0000FF">Dim</font> strFormData <font color="#0000FF">As
String</font><span style="mso-spacerun: yes"><br>
<br>
&nbsp;&nbsp;&nbsp; </span><font color="#008000">‘ Suppose that the version of
our app is “A”</font><span style="mso-spacerun: yes"><font color="#008000">&nbsp;</font>&nbsp;&nbsp;<br>
&nbsp;&nbsp;&nbsp; </span>strFormData = &quot;<b>Version=A&amp;Key=123ABC</b>&quot;<span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp; </span>strUrl = &quot;The address of the prev.. ASP
page&quot;</font></p>
<p class="MsoNormal"><font face="Courier New" size="2"><span style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;
</span>Inet1.url = strUrl<span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp; </span>Inet1.Execute strUrl, &quot;POST&quot;, strFormData, _<span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>&quot;Content-Type:
application/x-www-form-urlencoded&quot;<span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp; </span><font color="#008000">‘ …</font>&nbsp;<o:p>
</o:p>
</font></p>
<p class="MsoNormal"><font face="Courier New" size="2"><font color="#0000FF">Private
Sub </font>Inet1_StateChanged(<font color="#0000FF">ByVal</font> State <font color="#0000FF">As
Integer</font>)<span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp; </span><font color="#0000FF">Dim</font> strTemp <font color="#0000FF">As
String</font><span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp; </span><font color="#0000FF">If</font> <b>State</b> = <b>12</b>
<font color="#0000FF">Then</font> <font color="#008000">‘ If operation is
completed</font><span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>strTemp = <b>Inet1.GetChunk</b>(32000)<span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>strTemp = Trim(strTemp)<span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><font color="#0000FF">If</font>
strTemp = “to_app” <font color="#0000FF">then</font><span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>MsgBox
“You must update the application!”<span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><font color="#0000FF">ElseIf</font>
strTemp = “ok” <font color="#0000FF">then</font><span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>MsgBox
“No updates aviable”<span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><font color="#0000FF">Else</font><span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>MsgBox
“Unknow response”<span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span><font color="#0000FF">End If</font><span style="mso-spacerun: yes"><br>
&nbsp;&nbsp;&nbsp; </span><font color="#0000FF">End If</font><br>
<font color="#0000FF">End Sub</font></font><br>
<br>
<font face="Tahoma" size="2">As you can see this is only a silly little example,
but with this you can understand how to sending informations to a web-page in a
easy way.<br>
Remember that you must URL-Encode each byte you send to the page and that some
special chars may <b>lost</b> during encoding.<br>
<br>
That’s all, folks<br>
<br>
Hope this tutorial help you!<br>
SebaMix</font></p>

