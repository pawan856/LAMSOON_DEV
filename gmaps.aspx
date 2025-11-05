<%@ Page Language="VB" AutoEventWireup="false" CodeFile="gmaps.aspx.vb" Inherits="gmaps" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>TMS w/Google Maps</title>
    <link rel="stylesheet" href="stylesheet/general.css" type="text/css">
    <script src="http://maps.google.com/maps?file=api&v=2&sensor=false&key=ABQIAAAAQrBuAuJvIMyRCNL8Z0RcpRRDh-lWLqydV41b7Yr32yaEQsHZwRRSosBiW5UNyWOnzo4-lfj6ugW7uQ" type="text/javascript"></script>
    <script type="text/javascript" src="js/maps.streetviewcontrol.js"></script>
    <script type="text/javascript" language="javascript">
        var drawSatus = false;
        var isClickZoom = false;
        var iniPoint;
        var x_init=0; 
        var y_init=0; 
        var rectWidth=0; 
        var rectHeight=0; 
        var rectStatus = false; 


         //******************Custom control Definition*******************************
        //First, define the class
        function ClickZoomCtrl(){}
        //Derivate it
        ClickZoomCtrl.prototype = new GControl();
        
        //then, Initialize it
        ClickZoomCtrl.prototype.initialize = function(map){
          //Define map
          this.map=map;
          //Tool picture "dragZoomBtn" is contained into div element "container"
          var container = document.createElement("div");
          var ClickZoomBtn = document.createElement("img");
          container.appendChild(ClickZoomBtn);
          container.setAttribute("id","div_ClickZoomBtn");
          container.setAttribute("style","cursor:pointer;");
          container.setAttribute("title","Click twice & zoom");
          ClickZoomBtn.src = "http://www.googlemappers.com/libraries/magnifier/zoom2.png";
          //When clicking on it, activate click&Zoom function
          GEvent.addDomListener(container, "click", function(){
            clickZoomHandler(map);
          }
        );
    
          map.getContainer().appendChild(container);
          return container;
        }
    
        ClickZoomCtrl.prototype.getDefaultPosition=function()
        {
          return new GControlPosition(G_ANCHOR_TOP_LEFT,new GSize(683,8));
        }
        
        function clickZoomHandler(aMap){
          //activate click and zoom function
          isClickZoom = true;
          aMap.disableDragging();
          document.body.style.cursor = "crosshair";  
          //alert ("click and zoom activated");    
        }
        
        function initialize() 
        {
            if (!GBrowserIsCompatible()) return;
            
            var map = new GMap2(document.getElementById("map"));
            
            map.addControl(new GOverviewMapControl(new GSize(200,200)));
            map.addControl(new ClickZoomCtrl());
            map.addControl(new StreetViewControl);
            map.setCenter(new GLatLng(22.348806,114.159107), 11);
            map.setUIToDefault();
            GEvent.addListener(map, 'click', function(overlay, point){
            if (!isClickZoom) return;
            if (overlay) return; //overlay was clicked, do nothing : user must click once more time
                
            //click event has been raised for click&zoom operation
            //Check if it's 1st or 2nd click (begin or end of click&zoom operation)
              
            if (!drawSatus){
              //It's 1st click
              //store mousedown point coordinates into iniPoint
              iniPoint = point;
              
              //Set drawstatus to true
              drawSatus = true;
            }
            else{
              //It's 2nd click
              //!\ Don't set isClickZoom to false into this event, because it occurs 
                  //before selectrect click event, do it into this one !!
              //isClickZoom = false;
              drawSatus = false;
              //alert ("2nd click map");
                
              //Get correct SW and NE points coordinates
              if (iniPoint.lat()>point.lat()){
                if (iniPoint.lng()<point.lng()){
                  //maybe most case
                  var geoSW = new GLatLng(point.lat(),iniPoint.lng(),false);
                  var geoNE = new GLatLng(iniPoint.lat(),point.lng(),false);
                }
                else{
                  var geoSW = new GLatLng(point.lat(),point.lng(),false);
                  var geoNE = new GLatLng(iniPoint.lat(),iniPoint.lng(),false);
                }   
              }
              else{
                if (iniPoint.lng()<point.lng()){
                  var geoSW = new GLatLng(iniPoint.lat(),iniPoint.lng(),false);
                  var geoNE = new GLatLng(point.lat(),point.lng(),false);
                }
                else{
                  var geoSW = new GLatLng(iniPoint.lat(),point.lng(),false);
                  var geoNE = new GLatLng(point.lat(),iniPoint.lng(),false);
                }    
              }
        
              //define LatLngBound             
              var geoRectBounds = new GLatLngBounds(geoSW,geoNE);
              alert ("sw lat = "+ (geoRectBounds.getSouthWest()).lat()+"\n"+ "sw lng = "+ (geoRectBounds.getSouthWest()).lng()
                  +"\n"+"ne lat = "+ (geoRectBounds.getNorthEast()).lat()+"\n"+ "ne lng = "+ (geoRectBounds.getNorthEast()).lng());
                
              //Define new center
              var newCenter = new GLatLng(((geoRectBounds.getSouthWest()).lat()+(geoRectBounds.getNorthEast()).lat())/2,
                    ((geoRectBounds.getSouthWest()).lng()+(geoRectBounds.getNorthEast()).lng())/2 ,false );
              alert("new center= "+newCenter.lat()+" "+newCenter.lng()+"\n"+"current center= "+(map.getCenter()).lat()+" "+(map.getCenter()).lng());
                
              //zoom map !
              //var newZoom = map.getBoundsZoomLevel(geoRectBounds);
              //map.setCenter(newCenter,newZoom);
            }
    	    }
          );
     
          //******************Click event handler (on DOM Element)********************
          //Init
          //alert("init!!");
          document.getElementById("map").onmousemove = rectMouseMove;
          document.getElementById("map").onclick = rectClick;
          
          //Define rectangle element (div))
          var rect = document.createElement("div");
          rect.setAttribute("id","rectangle");
          //rect.setAttribute("class","rectangle");
          document.getElementById("map").appendChild(rect);
          
          function mouseX(evt) 
          {
            if (evt.pageX) return evt.pageX;
      	    else if (evt.clientX)
              return evt.clientX + (document.documentElement.scrollLeft ?
         	      document.documentElement.scrollLeft :
         	      document.body.scrollLeft);
      	    else return null;
          }
          
          function mouseY(evt) 
          {
      	    if (evt.pageY) return evt.pageY;
      	    else if (evt.clientY)
         	    return evt.clientY + (document.documentElement.scrollTop ?
         		    document.documentElement.scrollTop :
         		    document.body.scrollTop);
      	    else return null;
          }
          
          function rectClick(e) 
          {
            if (!e) var e = window.event;	// e gives access to the event in all browsers
          
            //Check status of rectangle
            if (!isClickZoom) return;
            
            if (!rectStatus)
            {
          
              //This is 1st cklick (rectangle is not being drawn)
              //alert ("1st click selectrect.js");
            
              //Store Mouse coordinates
              x_init=mouseX(e)- parseInt(document.getElementById("map").offsetLeft); 
              y_init=mouseY(e)- parseInt(document.getElementById("map").offsetTop); 
               
              //Rectangle Init 
              rectStatus=true;
              rectMouseMove(e);
              //Show it !
              document.getElementById("rectangle").style.visibility="visible";
              //alert (document.getElementById("rectangle").style.visibility);
            }
            else
            {
              //This is 2nd cklick (rectangle is being drawn)
              //alert ("2nd click rect");
              
              //Hide rectangle
              document.getElementById("rectangle").style.visibility="hidden"; 
              rectStatus=false;
              isClickZoom = false;
              map.enableDragging();
              document.body.style.cursor = "default";
            }
          
          }
          
          function rectMouseMove(e){
            
            if (!isClickZoom) return;
            if (!rectStatus) return;		// Not drawing rectangle
          
            if (!e) var e = window.event;	// e gives access to the event in all browsers
          
            //Store Mouse coordinates (+-5 : offset to capture click event on the map)
            var CurX = mouseX(e)- parseInt(document.getElementById("map").offsetLeft);
            var CurY = mouseY(e)- parseInt(document.getElementById("map").offsetTop);
            (CurX<x_init)? CurX +=5:CurX -=5;
            (CurY<y_init)? CurY +=5:CurY -=5;
                      
            //Store X and width
            rectWidth=Math.abs(CurX - x_init);
            if (CurX >= x_init) 
            { 
          
       	      document.getElementById('rectangle').style.left = x_init+"px"; 
            } 
            else 
            { 
       	      document.getElementById('rectangle').style.left = CurX+"px"; 
            } 
            document.getElementById('rectangle').style.width = rectWidth+"px";
            //alert ("mousemove4");
          
            //Store Y and height
            rectHeight=Math.abs(CurY - y_init); 
            if (CurY >= y_init) 
            { 
            document.getElementById("rectangle").style.top = y_init+"px"; 
            } 
            else 
            { 
              document.getElementById("rectangle").style.top = CurY+"px"; 
            } 
            document.getElementById("rectangle").style.height = rectHeight+"px";
            
          }
        }
    </script>
</head>
<body onload="javascript:initialize();" onunload="GUnload()">
    <form id="form1" runat="server">
        <div id="map" style="width: 1000px; height: 700px"></div>
    </form>
</body>
</html>
