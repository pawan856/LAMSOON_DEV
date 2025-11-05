
// Description:  Streetview InfoWindow GControl for Google Maps API V2.
// Author:       Marcel Wijnands
// Demo:         http://demo.shynet.nl/streetviewcontrol/
// Blog:         http://blog.shynet.nl/

function StreetViewControl() {
    this._map = null;
    this._client = null;
    this._panorama = null;
    this._marker = null;
    this._overlay = null;
    this._buttonOffsetRight = 215;
    this._buttonOffsetTop = 7;
    this._buttonGControlAnchor = G_ANCHOR_TOP_RIGHT;
    this._markerOffsetTop = 80;
} // StreetViewControl

StreetViewControl.prototype = new GControl();

StreetViewControl.prototype.initialize = function (map) {

    this._map = map;
    this._client = new GStreetviewClient();
    this._overlay = new GStreetviewOverlay();

    var guyIcon = new GIcon(G_DEFAULT_ICON);
    guyIcon.image = "http://maps.gstatic.com/mapfiles/cb/man_arrow-0.png";
    guyIcon.transparent = "http://maps.gstatic.com/mapfiles/cb/man-pick.png";
    guyIcon.imageMap = [26, 13, 30, 14, 32, 28, 27, 28, 28, 36, 18, 35, 18, 27, 16, 26, 16, 20, 16, 14, 19, 13, 22, 8];
    guyIcon.iconSize = new GSize(49, 52);
    guyIcon.iconAnchor = new GPoint(25, 35);
    guyIcon.infoWindowAnchor = new GPoint(25, 5);

    this._marker = new GMarker(new GLatLng(52.373198, 4.902948), { icon: guyIcon, draggable: true, hide: true, title: "Klik voor streetview weergave." });
    GEvent.bind(this._marker, "dragstart", this, this._onDragStart);
    GEvent.bind(this._marker, "dragend", this, this._onDragEnd);
    GEvent.bind(this._marker, "click", this, this._initPanorama);
    GEvent.bind(this._marker, "infowindowbeforeclose", this, this._removePanorama);
    GEvent.bind(this._marker, "infowindowopen", this, this._loadPanorama);
    this._map.addOverlay(this._marker);

    var containerDiv = document.createElement("div");
    containerDiv.title = "Streetview marker weergeven";
    containerDiv.style.backgroundColor = "white";
    containerDiv.style.border = "1px solid black";
    containerDiv.style.cursor = "pointer";
    containerDiv.style.textAlign = "center";
    containerDiv.style.width = "5em";
    var innerDiv = document.createElement("div");
    innerDiv.appendChild(document.createTextNode("Streetview"));
    innerDiv.style.borderBottom = "1px solid #b0b0b0";
    innerDiv.style.borderLeft = "1px solid white";
    innerDiv.style.borderRight = "1px solid #b0b0b0";
    innerDiv.style.borderTop = "1px solid white";
    innerDiv.style.color = "black";
    innerDiv.style.fontFamily = "Arial,sans-serif";
    innerDiv.style.fontSize = "12px";
    containerDiv.appendChild(innerDiv);

    GEvent.bindDom(containerDiv, "click", this, function () {
        if (this._marker.isHidden()) {
            this._marker.setLatLng(map.fromContainerPixelToLatLng(new GPoint((this._map.getContainer().offsetWidth - this._buttonOffsetRight) - (containerDiv.offsetWidth / 2), this._markerOffsetTop)));
            this._marker.show();
            innerDiv.style.borderBottom = "1px solid #6c9ddf";
            innerDiv.style.borderLeft = "1px solid #345684";
            innerDiv.style.borderRight = "1px solid #6c9ddf";
            innerDiv.style.borderTop = "1px solid #345684";
            innerDiv.style.fontWeight = "bold";
            containerDiv.title = "Streetview marker verbergen"
        } else {
            this._marker.hide();
            if (document.getElementById('StreetViewControl-Panorama')) {
                this._map.closeInfoWindow();
            }
            innerDiv.style.borderBottom = "1px solid #b0b0b0";
            innerDiv.style.borderLeft = "1px solid white";
            innerDiv.style.borderRight = "1px solid #b0b0b0";
            innerDiv.style.borderTop = "1px solid white";
            innerDiv.style.fontWeight = "normal";
            containerDiv.title = "Streetview marker weergeven"
        }
    });

    this._map.getContainer().appendChild(containerDiv);

    return containerDiv;

}  // initialize

StreetViewControl.prototype.getDefaultPosition = function () {
    return new GControlPosition(this._buttonGControlAnchor, new GSize(this._buttonOffsetRight, this._buttonOffsetTop));
} // getDefaultPosition

StreetViewControl.prototype._initPanorama = function (latlng) {
    var thisControl = this;
    thisControl._map.closeInfoWindow();
    thisControl._client.getNearestPanorama(latlng, function (data) {
        thisControl._initPanorama2.call(thisControl, data);
    });
}    // _initPanorama

StreetViewControl.prototype._initPanorama2 = function (data) {
    switch (data.code) {
        case 200:
            this._displayStreetInfoWindow.call(this, data.location.latlng);
            break;
        case 600:
            alert("Geen streetview beschikbaar op deze lokatie.");
            break;
        case 500:
            alert("De server reageert niet.");
            break;
    }
} // _initPanorama2

StreetViewControl.prototype._displayStreetInfoWindow = function (latlng) {
    var smallNode = document.createElement('div');
    smallNode.style.width = '400px';
    smallNode.style.margin = '15px';
    smallNode.style.height = '300px';
    smallNode.id = 'StreetViewControl-Panorama';
    smallNode.latlng = latlng;
    this._marker.openInfoWindow(smallNode);
} // _displayStreetInfoWindow

StreetViewControl.prototype._loadPanorama = function () {
    smallNode = document.getElementById('StreetViewControl-Panorama');
    this._panorama = new GStreetviewPanorama(smallNode);
    this._panorama.setLocationAndPOV(smallNode.latlng);
    GEvent.bind(this._panorama, "initialized", this, this._onNewLocation);
    GEvent.bind(this._panorama, "yawchanged", this, this._onYawChange);
} // _loadPanorama

StreetViewControl.prototype._removePanorama = function () {
    if (this._panorama) {
        this._panorama.remove();
    }
} // _removePanorama

StreetViewControl.prototype._onYawChange = function (newYaw) {
    var GUY_NUM_ICONS = 16;
    var GUY_ANGULAR_RES = 360 / GUY_NUM_ICONS;
    if (newYaw < 0) {
        newYaw += 360;
    }
    guyImageNum = Math.round(newYaw / GUY_ANGULAR_RES) % GUY_NUM_ICONS;
    guyImageUrl = "http://maps.gstatic.com/mapfiles/cb/man_arrow-" + guyImageNum + ".png";
    this._marker.setImage(guyImageUrl);
} // _onYawChange

StreetViewControl.prototype._onNewLocation = function (location) {
    this._marker.setLatLng(location.latlng);
} // _onNewLocation

StreetViewControl.prototype._onDragStart = function () {
    this._map.addOverlay(this._overlay);
} // _onDragStart

StreetViewControl.prototype._onDragEnd = function () {
    this._map.removeOverlay(this._overlay);
    if (document.getElementById('StreetViewControl-Panorama')) {
        this._initPanorama(this._marker.getLatLng());
    }
} // _onDragEnd

StreetViewControl.prototype._handleError = function (errorCode) {
    switch (errorCode) {
        case 600:
            alert("Geen streetview beschikbaar op deze lokatie.");
            break;
        case 602:
            alert("Geen gebruikersfoto gevonden.");
            break;
        case 603:
            alert("Geen flash plugin gedetecteerd.");
            break;
    }
} // _handleError
