import clr
clr.AddReference("System.Drawing")
clr.AddReference("Interop.SolidEdge")
clr.AddReference("System.Runtime.InteropServices")

import SolidEdgeFramework as SEFramework
import System.Runtime.InteropServices as SRI


class ExternalExample(object):
    _unitType = SEFramework.UnitTypeConstants.igUnitDistance
    _externalValue = "1 in"
    _internalValue = None

    def __init__(self):
        self.UpdateInternalValue()

    @property
    def UnitType(self):
        return self._unitType

    @UnitType.setter
    def UnitType(self, value):
        self._unitType = value
        self.UpdateInternalValue()
#
# 	<DescriptionAttribute("Value in user display units.")> _
    @property
    def ExternalValue(self):
        """Value in user display units."""
        return self._externalValue

    @ExternalValue.setter
    def ExternalValue(self, value):
        self._externalValue = value
        self.UpdateInternalValue()
#
# 	<DescriptionAttribute("Value in internal (database) units.")> _
    @property
    def InternalValue(self):
        """Value in internal (database) units."""
        return self._internalValue

    def UpdateInternalValue(self):
        try:
            application = SRI.Marshal.GetActiveObject("SolidEdge.Application")
            document = application.ActiveDocument
            unitsOfMeasure = document.UnitsOfMeasure
            self._internalValue = unitsOfMeasure.ParseUnit(self._unitType, self._externalValue)
        except Exception, ex:
            self._internalValue = ex.Message
        finally:
            if unitsOfMeasure is not None:
                SRI.Marshal.ReleaseComObject(unitsOfMeasure)
            if document is not None:
                SRI.Marshal.ReleaseComObject(document)
            if application is not None:
                SRI.Marshal.ReleaseComObject(application)
