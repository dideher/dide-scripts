from shapely.geometry import Point


'''
Searches a point inside all polygons (due to overlapping).
Creates a list of school names which correspond to this point.
Returns the list of school names.
'''


def searchPointInsidePolygons(point, namesList, polygonsList):
    result = list()

    for elem1, elem2 in zip(namesList, polygonsList):
        if point.within(elem2):
            result.append(elem1)

    return result


# For debugging: check if lists have the same size
def checkListsLengths(namesList, polygonsList):
    debugText = "Έλεγχος πλήθους ονομάτων και περιοχών Σχολικών Μονάδων\n"
    debugText += 60 * "-" + "\n"
    nls = len(namesList)
    debugText += "Πλήθος ονομάτων Σχολικών Μονάδων: {}\n".format(nls)
    pls = len(polygonsList)
    debugText += "Πλήθος περιοχών Σχολικών Μονάδων: {}\n".format(pls)

    sameSize = (nls == pls)

    return sameSize, debugText


# For debugging: check if polygons are valid
def checkPolygonsValidity(namesList, polygonsList):
    debugText = "\nΈλεγχος ορθότητας πολυγώνων περιοχών\n"
    debugText += 60 * "-" + "\n"
    polygonsAreValid = True

    for elem1, elem2 in zip(namesList, polygonsList):
        if not elem2.is_valid:
            polygonsAreValid = False
            debugText += "{}: δεν έχει οριστεί σωστά.\n".format(elem1)

    if polygonsAreValid:
        debugText += "Όλα τα πολύγωνα έχουν οριστεί σωστά.\n"

    return polygonsAreValid, debugText


'''
For debugging: C style code ;-)
Check for intersections between polygons.

An address can belong in more than one polygon, so you must check the whole polygonsList
(DO NOT break the loop when you have a match)
'''


def checkPolygonsIntersections(namesList, polygonsList):
    debugText = "\nΈλεγχος επικαλύψεων πολυγώνων περιοχών\n"
    debugText += 60 * "-" + "\n"
    counter = 0
    i = 0
    while i < len(namesList):
        y = i + 1
        while y < len(namesList):
            # if polygonsList[i].intersects(polygonsList[y]):
            if polygonsList[i].intersects(polygonsList[y]) and not polygonsList[i].touches(polygonsList[y]):
                counter += 1
                debugText += "{:3n}) Επικάλυψη μεταξύ: {} και {}\n".format(counter, namesList[i], namesList[y])
            y += 1
        i += 1

    noIntersections = (counter == 0)

    return noIntersections, debugText


# For debugging: check if known points are inside or outside
def checkPointInsideOutsidePolygon(polygonsList):
    pointInside = Point(35.323595, 25.135447)
    pointOutside = Point(35.340280, 25.141996)

    if pointInside.within(polygonsList[0]):
        print("pointInside: True")
    else:
        print("pointInside: False")

    if pointOutside.within(polygonsList[0]):
        print("pointOutside: True")
    else:
        print("pointOutside: False")
