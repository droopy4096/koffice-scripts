#!/usr/bin/env kross
# -*- coding: utf-8 -*-

"""
Python script to import content from a KPLato Project stored
within a KPlato file into KSpread.

(C)2008 Dag Andersen <danders@get2net.dk>
Licensed under LGPL v2+higher.
"""

import sys, os, traceback
# import Kross, KSpread
import Kross

try:
    import KSpread
except:
    KSpread=Kross.module("KSpread")

T = Kross.module("kdetranslation")

def i18n(text, args = []):
    if T is not None:
        return T.i18n(text, args)
    # No translation module, return the untranslated string
    for a in range( len(args) ):
        text = text.replace( ("%" + "%d" % ( a + 1 )), str(args[a]) )
    return text

def i18nc(comment, text, args = []):
    print args
    if T is not None:
        return T.i18n(comment, text, args)
    return i18n(text, args)

def i18np(singular, plural, number, args = []):
    if T is not None:
        return T.i18n(singular, plural, number, args)
    if number == 1:
        return i18n(singular, [number] + args)
    return i18n(plural, [number] + args)

def i18ncp(comment, singular, plural, number, args = []):
    if T is not None:
        return T.i18n(comment, singular, plural, number, args)
    return i18np(singular, plural, number, args)

class KPlatoImport:

    def __init__(self, scriptaction):
        self.scriptaction = scriptaction
        self.currentpath = self.scriptaction.currentPath()
        self.forms = Kross.module("forms")
        try:
            self.start()
        except Exception, inst:
            self.forms.showMessageBox("Sorry", i18n("Error"), "%s" % inst)
        except:
            self.forms.showMessageBox("Error", i18n("Error"), "%s" % "".join( traceback.format_exception(sys.exc_info()[0],sys.exc_info()[1],sys.exc_info()[2]) ))

    def start(self):
        writer = KSpread.writer()
        filename = self.showImportDialog(writer)
        if not filename:
            return # no exception, user prob pressed cancel

        KPlato = Kross.module("KPlato")
        if KPlato is None:
            raise Exception, i18n("Failed to start KPlato. Is KPlato installed?")

        KPlato.document().openUrl( filename )
        proj = KPlato.project()
        data = self.showDataSelectionDialog( writer, KPlato )
        if len(data) == 0:
            raise Exception, i18n("No data to import")

        objectType = data[0]
        schedule = data[1]
        props = data[2]
        if len(props) == 0:
            raise Exception, i18n("No properties to import")

        record = []
        if data[3] == True:
            for prop in props:
                record.append( KPlato.headerData( objectType, prop ) )
            if not writer.setValues(record):
                if self.forms.showMessageBox("WarningContinueCancel", i18n("Warning"), i18n("Failed to set all properties of '%1' to cell '%2'", [", ".join(record), writer.cell()])) == "Cancel":
                    return
            writer.next()

        if objectType == 0: # Nodes
            self.importValues( writer, KPlato, proj, props, schedule )
        if objectType == 1: # Resources
            for i in range( proj.resourceGroupCount() ):
                self.importValues( writer, KPlato, proj.resourceGroupAt( i ), props, schedule )
        if objectType == 2: # Accounts
            for i in range( proj.accountCount() ):
                self.importValues( writer, KPlato, proj.accountAt( i ), props, schedule )

    def importValues(self, writer, module, dataobject, props, schedule ):
        record = []
        for prop in props:
            record.append( module.data( dataobject, prop, "DisplayRole", schedule ) )
        if not writer.setValues(record):
            if self.forms.showMessageBox("WarningContinueCancel", i18n("Warning"), i18n("Failed to set all properties of '%1' to cell '%2'", [", ".join(record), writer.cell()])) == "Cancel":
                return
        writer.next()
        for i in range( dataobject.childCount() ):
            self.importValues( writer, module, dataobject.childAt( i ), props, schedule )

    def showImportDialog(self, writer):
        dialog = self.forms.createDialog(i18n("KPlato Import"))
        dialog.setButtons("Ok|Cancel")
        dialog.setFaceType("List") #Auto Plain List Tree Tabbed

        openpage = dialog.addPage(i18n("Open"),i18n("Import from KPlato Project File"),"document-import")
        openwidget = self.forms.createFileWidget(openpage, "kfiledialog:///kspreadkplatoimportopen")
        openwidget.setMode("Opening")
        openwidget.setFilter("*.kplato|%(1)s\n*|%(2)s" % { '1' : i18n("KPlato Project Files"), '2' : i18n("All Files") })
        if dialog.exec_loop():
            filename = openwidget.selectedFile()
            if not os.path.isfile(filename):
                raise Exception, i18n("No file selected.")
            return filename
        return None

    def showDataSelectionDialog(self, writer, KPlato ):
        tabledialog = self.forms.createDialog("Property List")
        tabledialog.setButtons("Ok|Cancel")
        tabledialog.setFaceType("List") #Auto Plain List Tree Tabbed
        
        datapage = tabledialog.addPage(i18n("Destination"),i18n("Import to sheet beginning at cell"))
        sheetslistview = KSpread.createSheetsListView(datapage)
        sheetslistview.setEditorType("Cell")

        schedulepage = tabledialog.addPage(i18n("Schedules"),i18n("Select schedule"))
        schedulewidget = KPlato.createScheduleListView(schedulepage)

        sourcepage = tabledialog.addPage(i18n("Data"),i18n("Select data"))
        sourcewidget = KPlato.createDataQueryView(sourcepage)
        if tabledialog.exec_loop():
            currentSheet = sheetslistview.sheet()
            if not currentSheet:
                raise Exception, i18n("No current sheet.")
            if not writer.setSheet(currentSheet):
                raise Exception, i18n("Invalid sheet '%1' defined.", [currentSheet])

            cell = sheetslistview.editor()
            if not writer.setCell(cell):
                raise Exception, i18n("Invalid cell '%1' defined.", [cell])

            schedule = schedulewidget.currentSchedule()
            #print "schedule: ", schedule
            props = sourcewidget.selectedProperties()
            #print "props: ", props
            ot = sourcewidget.objectType()
            #print "objectType: ", ot
            return [ot, schedule, props, sourcewidget.includeHeaders() ]
        return None


KPlatoImport( self )
