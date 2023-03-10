from shapely.geometry import Polygon
from pykml import parser


def createLists(filename):
    with open(filename, encoding='utf-8') as f:
        doc = parser.parse(f).getroot()

    namesList = list()
    coordsList = list()

    for placemark in doc.Document.iter('Placemark'):
        namesList.append(placemark.name.text)
        print(namesList[-1])
        coordsList.append(
            placemark.Polygon.outerBoundaryIs.LinearRing.coordinates.text.strip().replace(" ", "").replace(",0", ""))
        print(coordsList[-1])

    polygonsList = list()

    for elem1 in coordsList:
        pairsList = elem1.split("\n")
        pointsList = list()
        temp = list()

        for elem2 in pairsList:
            temp = elem2.split(",")
            point = (float(temp[1]), float(temp[0]))
            pointsList.append(point)

        polygonsList.append(Polygon(pointsList))

    return namesList, polygonsList
