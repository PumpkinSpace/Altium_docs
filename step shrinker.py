import OCC.BRep
import OCC.IFSelect
import OCC.Interface
import OCC.Quantity
import OCC.STEPCAFControl
import OCC.STEPControl
import OCC.TDataStd
import OCC.TCollection
import OCC.TColStd
import OCC.TDF
import OCC.TDocStd
import OCC.TopAbs
import OCC.TopoDS
import OCC.XCAFApp
import OCC.XCAFDoc
import OCC.XSControl

import logging

import aocutils.topology

import aocxchange.checks
import aocxchange.exceptions
import aocxchange.extensions

from aocxchange import step

import os

logger = logging.getLogger(__name__)


class StepXcafImporter(object):
    r"""Imports STEP file that support layers & colors"""
    def __init__(self, filename):

        aocxchange.checks.check_importer_filename(filename, aocxchange.extensions.step_extensions)

        self.filename = filename

        # The shape at index i in the following list corresponds
        # to the color and layer at index i in their respective lists
        self._shapes = list()
        self._colors = list()
        self._layers = list()

        self.read_file()

    @property
    def shapes(self):
        r"""Shapes"""
        return self._shapes

    @property
    def colors(self):
        r"""Colors"""
        return self._colors

    @property
    def layers(self):
        r"""Layers"""
        return self._layers

    @property
    def layers_str(self):
        r"""Returns a readable list of layers in the same order as self._shapes

        If self.shapes = [shape_1, shape_2], layers_str will return ['red', 'green'] when shape_1 is on the "red" layer
        and shape_2 is on the 'green' layer.

        See Also
        --------
        examples/export_multi_to_step_colors_layers_ocaf.py

        """
        layer_string_list = list()
        for i, layer in enumerate(self._layers):
            string = ""
            for j in range(1, layer.GetObject().Length() + 1):
                extended_string = layer.GetObject().Value(j)

                for k in range(1, extended_string.Length() + 1):
                    ascii_code = extended_string.Value(k)
                    string += (chr(ascii_code))

            layer_string_list.append(string)
        return layer_string_list

    def read_file(self):
        r"""Read file"""
        logger.info("Reading STEP file")
        h_doc = OCC.TDocStd.Handle_TDocStd_Document()

        # Create the application
        app = OCC.XCAFApp._XCAFApp.XCAFApp_Application_GetApplication().GetObject()
        app.NewDocument(OCC.TCollection.TCollection_ExtendedString("MDTV-CAF"), h_doc)

        # Get root assembly
        doc = h_doc.GetObject()
        h_shape_tool = OCC.XCAFDoc.XCAFDoc_DocumentTool().ShapeTool(doc.Main())
        color_tool = OCC.XCAFDoc.XCAFDoc_DocumentTool().ColorTool(doc.Main())
        layer_tool = OCC.XCAFDoc.XCAFDoc_DocumentTool().LayerTool(doc.Main())
        l_materials = OCC.XCAFDoc.XCAFDoc_DocumentTool().MaterialTool(doc.Main())

        step_reader = OCC.STEPCAFControl.STEPCAFControl_Reader()
        step_reader.SetColorMode(True)
        step_reader.SetLayerMode(True)
        step_reader.SetNameMode(True)
        step_reader.SetMatMode(True)

        status = step_reader.ReadFile(str(self.filename))

        if status == OCC.IFSelect.IFSelect_RetDone:
            logger.info("Transfer doc to STEPCAFControl_Reader")
            step_reader.Transfer(doc.GetHandle())

        labels = OCC.TDF.TDF_LabelSequence()
        color_labels = OCC.TDF.TDF_LabelSequence()
        # TopoDS_Shape a_shape;
        shape_tool = h_shape_tool.GetObject()
        shape_tool.GetShapes(labels)  # GetShapes was GetFreeShapes
        logger.info('Number of shapes at root :%i' % labels.Length())            
        color_tool.GetObject().GetColors(color_labels)
        logger.info('Number of colors : %i' % color_labels.Length())
        
        for i in range(labels.Length()):
            label = labels.Value(i + 1)
            print("")
            logger.debug("Label %i - type : %s" % (i+1, type(label)))
            logger.debug("Entry: %s" % label.EntryDumpToString())
            logger.debug("NbAttributes: %s" % label.NbAttributes())
            logger.debug("Is Assembly? %s" % shape_tool.IsAssembly(label))
            #logger.debug("Father %s" % label.Father())
            a_shape = shape_tool.GetShape(label)
            sub_shapes_labels = OCC.TDF.TDF_LabelSequence()
            has_subs = shape_tool.GetSubShapes(label, sub_shapes_labels)
            logger.debug('Has subshapes? %s' % has_subs)
            #logger.debug('Number of subshapes : %i' % sub_shapes_labels.Length())
            # Explore how to get part names
            h_name = OCC.TDataStd.Handle_TDataStd_Name()
            label.FindAttribute(OCC.TDataStd.TDataStd_Name_GetID(), h_name)
            name = OCC.TCollection.TCollection_ExtendedString()
            strname = name.PrintToString()
            logger.debug("Part name is: %s" % strname)
            logger.debug("Shape at label %i - type : %s" % (i+1, a_shape.ShapeType()))
            # string_seq = OCC.TColStd.TColStd_HSequenceOfExtendedString()
            # string_seq is an OCC.TColStd.TColStd_HSequenceOfExtendedString
            string_seq = layer_tool.GetObject().GetLayers(a_shape)
            color = OCC.Quantity.Quantity_Color()
            color_tool.GetObject().GetColor(a_shape, OCC.XCAFDoc.XCAFDoc_ColorSurf, color)
            if a_shape.ShapeType() == OCC.TopAbs.TopAbs_COMPOUND:
                logger.debug("Shape type: OCC.TopAbs.TopAbs_COMPOUND")
                topo = aocutils.topology.Topo(a_shape)
                logger.info("Nb of compounds : %i" % topo.number_of_compounds)
                logger.info("Nb of solids : %i" % topo.number_of_solids)
                logger.info("Nb of shells : %i" % topo.number_of_shells)
                for solid in topo.solids:
                    logger.info("Adding solid to the shapes list")
                    logger.info("Type : %s" % type(solid))
                    self._shapes.append(solid)
            elif a_shape.ShapeType() == OCC.TopAbs.TopAbs_SOLID:
                logger.debug("Shape type: OCC.TopAbs.TopAbs_SOLID")
                logger.debug("got general color: {0}, {1}, {2}".format(color.Red(), color.Green(), color.Blue()))
                self._colors.append(color)
                self._layers.append(string_seq)
        
        return True
    
    
step_import = StepXcafImporter('Z Sun Sensor Board.step')


out_filename = 'Z Sun Sensor Board2.step'

if not os.path.isfile(out_filename):
    os.makedirs('Z Sun Sensor Board2.step')
#end if

step_export = step.StepExporter(os.getcwd() + '\\' + out_filename)

shape_list = []

for shape in step_import.shapes:
    new_shape = True
    for old_shape in shape_list:
        if shape.IsPartner(old_shape):
            new_shape = False
            print "This shape is a duplicate"
        #end if
    #end for
    
    if new_shape:
        shape_list.append(shape)
        step_export.add_shape(shape)
    #end if
#end for

step_export.write_file()
