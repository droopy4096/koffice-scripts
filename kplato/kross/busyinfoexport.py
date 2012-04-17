#!/usr/bin/env kross
# -*- coding: utf-8 -*-

from __future__ import print_function
import os, sys, traceback
import Kross

from icalendar import Calendar, Event

try:
    import KPlato
    cli=False
except:
    KPlato=Kross.module("KPlato")
    cli=True

# from icalendar import UTC # timezone
# from icalendar import Calendar, Event
from datetime import datetime

T = Kross.module("kdetranslation")

def i18n(text, args = []):
    if T is not None:
        return T.i18n(text, args).decode('utf-8')
    # No translation module, return the untranslated string
    for a in range( len(args) ):
        text = text.replace( ("%" + "%d" % ( a + 1 )), str(args[a]) )
    return text

class ICSBusyinfoExporter:

    def __init__(self, scriptaction,cli=False):
        self.scriptaction = scriptaction
        self.currentpath = self.scriptaction.currentPath()

        self.cli=cli

        self.forms = Kross.module("forms")

        if not self.cli:
            # launched from inside KPlato
            self.proj = KPlato.project()
        else:
            # launched as a standalone script...
            print("opening doc...")
            filename=self.showImportDialog()
            KPlato.document().openUrl(filename)
            print( filename)
            self.proj=KPlato.project()
        
        self.dialog = self.forms.createDialog(i18n("Busy Information Export"))
        self.dialog.setButtons("Ok|Cancel")
        self.dialog.setFaceType("List") #Auto Plain List Tree Tabbed

        datapage = self.dialog.addPage(i18n("Schedules"),i18n("Export Selected Schedule"),"document-export")
        self.scheduleview = KPlato.createScheduleListView(datapage)
        
        savepage = self.dialog.addPage(i18n("Save"),i18n("Export Busy Info File"),"document-save")
        self.savewidget = self.forms.createFileWidget(savepage, "kfiledialog:///kplatobusyinfoexportsave")
        self.savewidget.setMode("Saving")
        self.savewidget.setFilter("*.ics|%(1)s\n*|%(2)s" % { '1' : i18n("Calendar"), '2' : i18n("All Files") } )

        if self.dialog.exec_loop():
            try:
                self.doExport( self.proj )
            except:
                self.forms.showMessageBox("Error", i18n("Error"), "%s" % "".join( traceback.format_exception(sys.exc_info()[0],sys.exc_info()[1],sys.exc_info()[2]) ))



    ### cal = Calendar()

    ### event = Event()
    ### event.add('summary', 'Python meeting about calendaring')
    ### event.add('dtstart', datetime(2005,4,4,8,0,0,tzinfo=UTC))
    ### event.add('dtend', datetime(2005,4,4,10,0,0,tzinfo=UTC))
    ### event.add('dtstamp', datetime(2005,4,4,0,10,0,tzinfo=UTC))
    ### event['uid'] = '20050115T101010/27346262376@mxm.dk'
    ### event.add('priority', 5)

    ### cal.add_component(event)

    ### f = open('example.ics', 'wb')
    ### f.write(cal.as_string())
    ### f.close()

    def doExport( self, project ):
        filename = self.savewidget.selectedFile()
        if not filename:
            self.forms.showMessageBox("Sorry", i18n("Error"), i18n("No file selected"))
            return
        schId = self.scheduleview.currentSchedule()
        if schId == -1:
            self.forms.showMessageBox("Sorry", i18n("Error"), i18n("No schedule selected"))
            return
        # cal= Calendar()
        # cal.add('prodid', '-//'+project.id()+'//'+KPlato.data(project, 'NodeName'))
        # cal.add('version', '0.1')
        file = open( filename, 'wb' )
        print("=================== "+str(project.id())+str(KPlato.data(project,'NodeName'))+" ==================",file=file)
        for i in range( project.resourceGroupCount() ):
            g = project.resourceGroupAt( i )
            print (g,file=file)
            for ri in range( g.resourceCount() ):
                r = g.resourceAt( ri )
                print("====> r:\n\t", "\n\t".join( dir(r)),file=file)
                lst = r.appointmentIntervals( schId )
                print("====> lst:\n\t", "\n\t".join(dir(lst)),file=file)
                for iv in lst:
                    print(r.id(), KPlato.data( r, 'ResourceName' ), iv, sep=" | ", file=file)
                    print ("====> iv:\n\t", "\n\t".join( dir(iv)),file=file)
                    # iv.insert( 0, r.id() )
                    # iv.insert( 1, KPlato.data( r, 'ResourceName' ) )
                    # pickle.dump( iv, file )

        file.close()

    def showImportDialog(self):
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

ICSBusyinfoExporter( self , cli=cli)
