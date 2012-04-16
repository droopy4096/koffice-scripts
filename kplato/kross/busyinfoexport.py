#!/usr/bin/env kross
# -*- coding: utf-8 -*-

import os, sys, traceback
import Kross, KPlato
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

    def __init__(self, scriptaction):
        self.scriptaction = scriptaction
        self.currentpath = self.scriptaction.currentPath()

        self.proj = KPlato.project()
        
        self.forms = Kross.module("forms")
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
        print "=================== foo =================="
        # cal= Calendar()
        # cal.add('prodid', '-//'+project.id()+'//'+KPlato.data(project, 'NodeName'))
        # cal.add('version', '0.1')
        file = open( filename, 'wb' )
        for i in range( project.resourceGroupCount() ):
            g = project.resourceGroupAt( i )
            print g
            for ri in range( g.resourceCount() ):
                r = g.resourceAt( ri )
                lst = r.appointmentIntervals( schId )
                for iv in lst:
                    print r.id(), KPlato.data( r, 'ResourceName' ), iv
                    # iv.insert( 0, r.id() )
                    # iv.insert( 1, KPlato.data( r, 'ResourceName' ) )
                    # pickle.dump( iv, file )

        file.close()

ICSBusyinfoExporter( self )
