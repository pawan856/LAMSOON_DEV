function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

var jsDateFormat = "YYYY/MM/DD";
//var jsDateFormat = "DD/MM/YYYY";


//------ Date Picker ------
nombresMes = Array("","January","February","March","April","May","June","July","August","September","October","November","December");

var anoInicial = 1900;
var anoFinal = 2100;
var ano;
var mes;
var dia;
var changeDateFlag;
var YFlag;
var campoDeRetorno;
var titulo;

YFlag = 'Y';

function diasDelMes(ano,mes) {
       if ((mes==1)||(mes==3)||(mes==5)||(mes==7)||(mes==8)||(mes==10)||(mes==12)) dias=31
  else if ((mes==4)||(mes==6)||(mes==9)||(mes==11)) dias=30
  else if ((((ano % 100)==0) && ((ano % 400)==0)) || (((ano % 100)!=0) && ((ano % 4)==0))) dias = 29
  else dias = 28;
  return dias;
};

function crearSelectorMes(mesActual) {
  var selectorMes = "", showMonth;
  selectorMes = "<select name='mes' size='1' onChange='opener.dibujarMes(self.document.Forma1.ano[self.document.Forma1.ano.selectedIndex].value,self.document.Forma1.mes[self.document.Forma1.mes.selectedIndex].value);'>\r\n";
  for (var i=1; i<=12; i++) {
    if (i < 10) {
        showMonth = "0" + i;
    } else {
        showMonth = i;
    }
    selectorMes = selectorMes + "  <option value='" + showMonth + "'";
    if (i == mesActual) selectorMes = selectorMes + " selected";
    selectorMes = selectorMes + ">" + nombresMes[i] + "</option>\r\n";
  }
  selectorMes = selectorMes + "</select>\r\n";
  return selectorMes;
}

function crearSelectorAno(anoActual) {
  var selectorAno = "";
  selectorAno = "<select name='ano' size='1' onChange='opener.dibujarMes(self.document.Forma1.ano[self.document.Forma1.ano.selectedIndex].value,self.document.Forma1.mes[self.document.Forma1.mes.selectedIndex].value);'>\r\n";
  for (var i=anoInicial; i<=anoFinal; i++) {
    selectorAno = selectorAno + "  <option value='" + i + "'";
    if (i == anoActual) selectorAno = selectorAno + " selected";
    selectorAno = selectorAno + ">" + i + "</option>\r\n";
  }
  selectorAno = selectorAno + "</select>";
  return selectorAno;
}

function crearTablaDias(numeroAno,numeroMes) {
  var tabla = "<table border='0' cellpadding='2' cellspacing='0' bgcolor='#EAEBE6'>\r\n  <tr>";
  var fechaInicio = new Date();
  fechaInicio.setDate(1);
  fechaInicio.setYear(numeroAno);
  fechaInicio.setMonth(numeroMes-1);
  ajuste = fechaInicio.getDay();
  tabla = tabla + "\r\n    <td align='center'>Su</td><td align='center'>Mo</td><td align='center'>Tu</td><td align='center'>We</td><td align='center'>Th</td><td align='center'>Fr</td><td align='center'>Sa</td></div>\r\n  <tr>";
  for (var j=1; j<=ajuste; j++) {
    tabla = tabla + "\r\n    <td></td>";
  }
  for (var i=1; i<10; i++) {
    tabla = tabla + "\r\n    <td"
    if ((i == diaHoy()) && (numeroMes == mesHoy()) && (numeroAno == anoHoy())) tabla = tabla + " bgcolor='#F7F7E7'";
    tabla = tabla + "><input type='button' style='background:#1F5AA6; color=#FFFFFF' value='0" + i + "' onClick='opener.changeDateFlag=opener.YFlag;opener.ano=self.document.Forma1.ano[self.document.Forma1.ano.selectedIndex].value; opener.mes=self.document.Forma1.mes[self.document.Forma1.mes.selectedIndex].value; opener.dia=" + i + "; self.close();' id='button'1 name='button'1></td>";
    if (((i+ajuste) % 7)==0) tabla = tabla + "\r\n  </tr>\r\n\  <tr>";
  }
  for (var i=10; i<=diasDelMes(numeroAno,numeroMes); i++) {
    tabla = tabla + "\r\n    <td"
    if ((i == diaHoy()) && (numeroMes == mesHoy()) && (numeroAno == anoHoy())) tabla = tabla + " bgcolor='#F7F7E7'";
    tabla = tabla + "><input type='button' style='background:#1F5AA6; color=#FFFFFF' value='" + i + "' onClick='opener.changeDateFlag=opener.YFlag;opener.ano=self.document.Forma1.ano[self.document.Forma1.ano.selectedIndex].value; opener.mes=self.document.Forma1.mes[self.document.Forma1.mes.selectedIndex].value; opener.dia=" + i + "; self.close();' id='button'1 name='button'1></td>";
    if (((i+ajuste) % 7)==0) tabla = tabla + "\r\n  </tr>\r\n\  <tr>";
  }
  tabla = tabla + "\r\n  </tr>\r\n</table>";
  return tabla;
}

function dibujarMes(numeroAno,numeroMes) {

  var html = "";
  var mes1 = "";
  html = html + "<html>\r\n<head>\r\n  <title>" + titulo + "</title>\r\n</head>\r\n<body bgcolor='#EAEBE6' onUnload='opener.escribirFecha();'>\r\n  <div align='center'>\r\n  <form name='Forma1'>\r\n";
  html = html + crearSelectorMes(numeroMes);
  html = html + crearSelectorAno(numeroAno);
  html = html + crearTablaDias(numeroAno,numeroMes);
  if (mes < 10) 
  {
	mes1 = "0" + mes;
  }
  else
  {
	mes1 = mes;
  }
  html = html + "<center><p><input type='button' style='background:#6699CC' name='hoy' value='Today: " + javaDateConverter("MM/DD/YYYY",jsDateFormat,mes1 + "/" + dia + "/" + ano) + "' onClick='opener.changeDateFlag=opener.YFlag;opener.ano="+ano+";opener.mes="+mes1+";opener.dia="+dia+"; self.close();'></center>";
  html = html + "\r\n  </form>\r\n  </div>\r\n</body>\r\n</html>\r\n";
  ventana = open("","calendario","width=220,height=270,top=200,left=300");
  ventana.document.open();
  ventana.document.writeln(html);
  ventana.document.close();
  ventana.focus();
}

function anoHoy() {
  var fecha = new Date();
  if (navigator.appName == "Netscape") return fecha.getYear() + 1900
  else return fecha.getYear();
}

function mesHoy() {
  var fecha = new Date();
  return fecha.getMonth()+1;
}

function diaHoy() {
  var fecha = new Date();
  return fecha.getDate();
}

function pedirFecha(campoTexto,nombreCampo) {
  ano = anoHoy();
  mes = mesHoy();
  dia = diaHoy();
  changeDateFlag = "";
  campoDeRetorno = campoTexto;
  titulo = nombreCampo;
  if(campoDeRetorno.disabled==false)
    dibujarMes(ano,mes);
}

function escribirFecha() {
	if (changeDateFlag=='Y')
	{
	  //campoDeRetorno.value = mes + "/" + dia + "/" + ano ;
	  campoDeRetorno.value = javaDateConverter("MM/DD/YYYY",jsDateFormat,mes + "/" + dia + "/" + ano);

	  campoDeRetorno.focus();
	}  
}


	function javaDateConverter(from_date,to_date,date_content)
	{

		var fDateFormat = from_date;
		var tDateFormat = to_date;
		var xDate = date_content;
		var final_date;
		var dateMonth;
		var dateDay;
		var dateYear;
		var datefullYear;

/*		
		//check for valid input

		if (fDateFormat.length != xDate.length)
		{
			alert("Error: Invalid Input Data");
			return false;
		}
*/				
		splitDateStr = "/"	
		var fdateArray = fDateFormat.split(splitDateStr);
		var xDateArray = xDate.split(splitDateStr);
		var tdateArray = tDateFormat.split(splitDateStr);		
		
		for (i = 0; i <= 2; i++)
		{
			switch(fdateArray[i])
			{ 
				case "MM":
					if (checkMM(xDateArray[i]) == false)
					{
						alert("Please input the valid month" + xDateArray[i]);
						return false;
					}
					else
					{
						dateMonth = xDateArray[i];
						
						if (dateMonth.length == 1)
						{
							dateMonth = "0" + dateMonth
						}  
					}
				break;
				case "DD":
					if (checkDD(xDateArray[i]) == false)
					{
						alert("Please input the valid day");
						return false;
					}					
					else
					{
						dateDay = xDateArray[i];
						
						if (dateDay.length == 1)
						{
							dateDay = "0" + dateDay
						}  
					}
				break;
				case "YYYY":
					if (checkYYYY(xDateArray[i]) == false)
					{
						alert("Please input the valid year");
						return false;
					}			
					else
					{
						dateFullYear = xDateArray[i];
						dateYear = dateFullYear.substr(2,2);				
					}
				break;					
				case "YY":
					if (checkyy(xDateArray[i]) == true)
					{
						dateYear =xDateArray[i]; 
						dateFullYear = returnFullYear(xDateArray[i]);
					} 
					else
					{
						alert("Please input the valid year");
						return false;
					}
			}
		}
		final_date = "";
		
		for (k = 0; k <= 2; k++)
		{ 
			switch(tdateArray[k])
			{
			case "MM":
				final_date = final_date + dateMonth + splitDateStr;
			    break;
			case "DD":
				final_date = final_date + dateDay + splitDateStr;
                break;
			case "YYYY":
				final_date = final_date + dateFullYear + splitDateStr;
			    break;
			case "YY":
				final_date = final_date + dateYear + splitDateStr;
			}
		}

        final_date = final_date.substring(0, final_date.length - 1);    

		if (checkdate(dateDay, dateMonth, dateYear) == false)
		{
			alert("Please enter a valid date");
			return false;
		}
		else
		{
			return final_date;
		}		
	}		
		
		
//Month validation	
function checkMM(mm)
{			
	if (isInteger(mm))
	{			
		if ((mm >= 1) && (mm <= 12 ))
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	else
	{
		return false;
	}
}
		
//Day validation
function checkDD(dd)
{
	if (isInteger(dd))
	{
		return true;
	}
	else
	{
		return false;
	}
}		
		
//FullYear validation
function checkYYYY(yyyy)			
{
	if (isInteger(yyyy) && yyyy.length == 4)
	{
		return true;
	}
	else
	{	
		return false;
	}
}
		
//return FullYear
function returnFullYear(yy)		
{
	if (parseInt(yy) > 50)
	{
		return "19" + yy;
	}
	else
	{
		return "20" + yy;
	}
}

//Year validation
function checkyy(yy)
{
	if (isInteger(yy) && yy.length == 2)
	{
		return true;
	}
	else
	{
		return false;
	}
}								
				
function checkdate(d,m,y) 
{	
	 if (m<1 || m>12)  return(false); 
	 if (d<1 || d>31)  return(false); 
	 if (m==4 || m==6 || m==9 || m==11) 
	   if (d==31) return(false); 
	 if (m==2) 
	 { 
	   var b=parseInt(y/4); 
	   if (isNaN(b)) return(false); 
	   if (d>29)     return(false); 
	   if (d==29 && ((y/4)!=parseInt(y/4))) return(false); 
	 } 
	 return(true); 
}