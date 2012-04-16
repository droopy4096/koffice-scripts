#!/usr/bin/env kross
# -*- coding: utf-8 -*-

import os, sys, traceback, tempfile, zipfile
import Kross
import KPlato

T = Kross.module("kdetranslation")

State_Started = 0
State_StartedLate = 1
State_StartedEarly = 2
State_Finished = 3
State_FinishedLate = 4
State_FinishedEarly = 5
State_Running = 6
State_RunningLate = 7
State_RunningEarly = 8
State_ReadyToStart = 9        # all precceeding tasks finished (if any)
State_NotReadyToStart = 10    # all precceeding tasks not finished (must be one or more)
State_NotScheduled = 11

def i18n(text, args = []):
    if T is not None:
        return T.i18n(text, args).decode('utf-8')
    # No translation module, return the untranslated string
    for a in range( len(args) ):
        text = text.replace( ("%" + "%d" % ( a + 1 )), str(args[a]) )
    return text

class BusyinfoExporter:

    def __init__(self, scriptaction):
        self.scriptaction = scriptaction
        self.currentpath = self.scriptaction.currentPath()

        self.proj = KPlato.project()
        
        self.forms = Kross.module("forms")
        self.dialog = self.forms.createDialog(i18n("Project Information Export"))
        self.dialog.setButtons("Ok|Cancel")
        self.dialog.setFaceType("List") #Auto Plain List Tree Tabbed

        datapage = self.dialog.addPage(i18n("Schedules"),i18n("Export Selected Schedule"),"document-export")
        self.scheduleview = KPlato.createScheduleListView(datapage)
        
        savepage = self.dialog.addPage(i18n("Save"),i18n("Export Project Info File"),"document-save")
        self.savewidget = self.forms.createFileWidget(savepage, "kfiledialog:///kplatoprojinfoexportsave")
        self.savewidget.setMode("Saving")
        self.savewidget.setFilter("*.txt|%(1)s\n*|%(2)s" % { '1' : i18n("Text Files"), '2' : i18n("All Files") } )

        if self.dialog.exec_loop():
            try:
                self.doExport( self.proj )
            except:
                self.forms.showMessageBox("Error", i18n("Error"), "%s" % "".join( traceback.format_exception(sys.exc_info()[0],sys.exc_info()[1],sys.exc_info()[2]) ))

    def doExport( self, project ):
        filename = self.savewidget.selectedFile()
        if not filename:
            self.forms.showMessageBox("Sorry", i18n("Error"), i18n("No file selected"))
            return
        schId = self.scheduleview.currentSchedule()
        if schId == -1:
            self.forms.showMessageBox("Sorry", i18n("Error"), i18n("No schedule selected"))
            return
        p = []

        ## p.append( project.id() )
        ## p.append( KPlato.data( project, 'NodeName' ) )
        ## pickle.dump( p, file )
        ## for i in range( project.resourceGroupCount() ):
            ## g = project.resourceGroupAt( i )
            ## for ri in range( g.resourceCount() ):
                ## r = g.resourceAt( ri )
                ## lst = r.appointmentIntervals( schId )
                ## for iv in lst:
                    ## iv.insert( 0, r.id() )
                    ## iv.insert( 1, KPlato.data( r, 'ResourceName' ) )
                    ## pickle.dump( iv, file )


        file = open( filename, 'wb' )
        file.write("\n".join(self.get_tasks(schId)))
        file.close()

    def get_tasks(self,sid):
        def printNodes( node, props, schedule, types = None ):
            ret=[]
            ret.append( printNode( node, props, schedule, types ) )
            for i in range( node.childCount() ):
                ret=ret+printNodes( node.childAt( i ), props, schedule, types )
            return ret

        def printNode( node, props, schedule, types = None ):
            s=''
            if types is None or node.type() in types:
                for prop in props:
                    s=s+"%-25s" % ( KPlato.data( node, prop[0], prop[1], schedule ) )
            return s

        ret=[]
        nodeprops = [['NodeWBSCode', 'DisplayRole'], ['NodeName', 'DisplayRole'], ['NodeType', 'DisplayRole'], ['NodeResponsible', 'DisplayRole'], ['NodeStatus', 'EditRole'] ]
        ret.append("Print tasks and milestones in arbitrary order:")
        # print the localized headers
        s=''
        for prop in nodeprops:
            s1="%-25s" % (self.proj.nodeHeaderData( prop ) )
            s=s+s1
        ret.append(s)
        ret.append('')
        ret=ret+printNodes( self.proj, nodeprops, sid, [ 'Task', 'Milestone' ] )
        ret.append('')
        return ret

BusyinfoExporter( self )


