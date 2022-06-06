from shapely.geometry import Point
from geopy.geocoders import GoogleV3, Here, Bing
import geocoder
import googlemaps


def searchAddressInGoogleMaps(myKey, searchItem):
    try:
        gmaps = googlemaps.Client(key=myKey)

        # Geocoding an address
        heraklionBounds = {"northeast": {"lat": 35.4669557, "lng": 25.5504097}, "southwest": {"lat": 34.9194342, "lng": 24.7224739}}
        result = gmaps.geocode(address=searchItem, bounds=heraklionBounds, language="el")
    except:
        return Point(0, 0), "Exception."
    else:
        if result:
            point = Point(result[0]['geometry']['location']['lat'], result[0]['geometry']['location']['lng'])
            addressFromPoint = result[0]['formatted_address']
        else:
            return Point(0, 0), "Can't find the address."

    return point, addressFromPoint


def searchAddressInGoogleV3(myKey, searchItem):
    try:
        g = GoogleV3(api_key=myKey, timeout=1000)
        result = g.geocode(searchItem, exactly_one=True, language="el")
    except:
        return Point(0, 0), "Exception."
    else:
        if result:
            point = Point(result.latitude, result.longitude)
            addressFromPoint = result.address
        else:
            point = Point(0, 0)
            addressFromPoint = "Can't find the address."

    return point, addressFromPoint


def searchAddressInBingMaps(myKey, searchItem):
    try:
        result = geocoder.bing(searchItem, adminDistrict='Heraklion', method='details', key=myKey)
    except:
        return Point(0, 0), "Exception."
    else:
        if result:
            point = Point(result.latlng[0], result.latlng[1])
            addressFromPoint = result.address
        else:
            point = Point(0, 0)
            addressFromPoint = "Can't find the address."

    return point, addressFromPoint


def searchAddressInHereMaps(myKey, searchItem):
    try:
        here = Here(apikey=myKey, timeout=1000)
        result = here.geocode(searchItem, exactly_one=True, language="el")
    except:
        return Point(0, 0), "Exception."
    else:
        if result:
            point = Point(result.latitude, result.longitude)
            addressFromPoint = result.address
        else:
            point = Point(0, 0)
            addressFromPoint = "Can't find the address."

    return point, addressFromPoint
