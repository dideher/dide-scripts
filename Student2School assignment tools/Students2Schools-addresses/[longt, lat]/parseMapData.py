from shapely.geometry import Polygon
import xml.etree.ElementTree as ET


'''
IMPORTANT:
The mapdata file must be exported only with the areas information (otherwise it creates folders).
You have to delete the namespace from the <kml> tag, otherwise it doesn't work

Reads the map data file, creates and returns two lists:
1) the list with school names
2) the list with school polygons
'''
def createLists(filename):
    tree = ET.parse(filename)
    root = tree.getroot()

    namesList = list()
    coordsList = list()

    for placemark in root.findall("./Document/Placemark"):
        namesList.append(placemark.find("name").text)
        # print(placemark.find("name").text)
        if placemark.find("./Polygon/outerBoundaryIs/LinearRing/coordinates") is not None:
            coordsList.append(placemark.find("./Polygon/outerBoundaryIs/LinearRing/coordinates").text.strip().replace(" ", "").replace(",0", ""))
            # print(placemark.find("./Polygon/outerBoundaryIs/LinearRing/coordinates").text)
        else:
            coordsList.append(placemark.findall("./MultiGeometry/Polygon/outerBoundaryIs/LinearRing/coordinates")[1].text.strip().replace(" ", "").replace(",0", ""))
            # print(placemark.findall("./MultiGeometry/Polygon/outerBoundaryIs/LinearRing/coordinates")[1].text)

    polygonsList = list()

    for elem1 in coordsList:
        pairsList = elem1.split("\n")
        pointsList = list()
        temp = list()

        for elem2 in pairsList:
            temp = elem2.split(",")
            point = (float(temp[0]), float(temp[1]))
            pointsList.append(point)

        polygonsList.append(Polygon(pointsList))

    return namesList, polygonsList