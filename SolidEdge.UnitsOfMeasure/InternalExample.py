import clr
clr.AddReference("System.Drawing")
clr.AddReference("Interop.SolidEdge")
clr.AddReference("System.Runtime.InteropServices")

import SolidEdgeFramework as SEFramework
import SolidEdgeConstants as SEConstants
import System.Runtime.InteropServices as SRI

class InternalExample(object):
    _unitType = SEFramework.UnitTypeConstants.igUnitDistance
    _precision = SEConstants.PrecisionConstants.igPrecisionThousandths
    _externalValue = None
    _internalValue = 0.0254

    def __init__(self):
        self.UpdateExternalValue()

    @property
    def UnitType(self):
       return self._unitType

    @UnitType.setter
    def UnitType(self, value):
        self._unitType = value
        self.UpdateExternalValue()

    @property
    def Precision(self):
       return self._precision

    @Precision.setter
    def Precision(self, value):
        self._precision = value
        self.UpdateExternalValue()

#<DescriptionAttribute("Value in user display units.")> _
    @property
    def ExternalValue(self):
        """Value in user display units."""
        return self._externalValue

#<DescriptionAttribute("Value in internal (database) units.")> _
    @property
    def InternalValue(self):
        """Value in internal (database) units."""
        return self._internalValue

    @InternalValue.setter
    def InternalValue(self, value):
        self._internalValue = value
        self.UpdateExternalValue()

    def UpdateExternalValue(self):
        try:
            application = SRI.Marshal.GetActiveObject("SolidEdge.Application")
            document = application.ActiveDocument
            unitsOfMeasure = document.UnitsOfMeasure
            self._externalValue = unitsOfMeasure.FormatUnit(self._unitType, self._internalValue, self._precision)
        except Exception, ex:
            self._externalValue = ex.Message
        finally:
            if unitsOfMeasure is not None:
                SRI.Marshal.ReleaseComObject(unitsOfMeasure)
            if document is not None:
                SRI.Marshal.ReleaseComObject(document)
            if application is not None:
                SRI.Marshal.ReleaseComObject(application)
