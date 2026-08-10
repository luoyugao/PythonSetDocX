wdHeaderFooterPrimary = 1
wdHeaderFooterFirstPage = 2
wdHeaderFooterEvenPages = 3

wdAlignParagraphLeft = 0
wdAlignParagraphCenter = 1
wdAlignParagraphRight = 2
wdAlignParagraphJustify = 3
wdAlignParagraphDistribute = 4

wdAlignPageNumberLeft = 1
wdAlignPageNumberCenter = 2
wdAlignPageNumberRight = 3

wdPageNumberStyleArabic = 1

wdFieldPage = 33
wdFieldNumPages = 26

wdNoProtection = 0

wdStatisticPages = 2

wdLineSpaceSingle = 0
wdLineSpace1pt5 = 1
wdLineSpaceDouble = 2
wdLineSpaceAtLeast = 3
wdLineSpaceExactly = 4
wdLineSpaceMultiple = 5

wdInlineShapePicture = 3

wdStyleTypeParagraph = 1

wdOutlineNumberGallery = 2
wdListNumberStyleArabic = 0

wdWrapSquare = 0
wdWrapTopBottom = 3

wdAutoFitWindow = 2

msoAutomationSecurityLow = 1

wdFirstCharacterLineNumber = 10
wdWithInTable = 12
wdRelativeHorizontalPositionPage = 0
wdShapeCenter = -999994

wdLineStyleSingle = 1
wdLineWidth025pt = 2
wdLineWidth075pt = 8
wdColorBlack = 0

wdBorderTop = -1
wdBorderLeft = -2
wdBorderBottom = -3
wdBorderRight = -4

wdListNumberStyleSimpChinNum = 37
wdListNumberStyleNumberInCircle = 18
wdListNumberStyleNone = 255

wdTrailingTab = 0
wdTrailingSpace = 1
wdTrailingNone = 2


def set_range_style(range_obj, style):
    try:
        range_obj.set_Style(style)
    except AttributeError:
        range_obj.Style = style


def get_range_information(range_obj, info_type):
    try:
        return range_obj.Information[info_type]
    except TypeError:
        return range_obj.Information(info_type)


def get_range_style(range_obj):
    try:
        return range_obj.get_Style()
    except AttributeError:
        return range_obj.Style


def get_list_gallery(app_obj, gallery_type):
    try:
        return app_obj.ListGalleries[gallery_type]
    except TypeError:
        return app_obj.ListGalleries(gallery_type)