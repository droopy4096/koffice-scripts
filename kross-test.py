#!/usr/bin/env kross
# -*- coding: utf-8 -*-

import Kross

try:
    import KSpread
except:
    KSpread=Kross.module("KSpread")

try:
    import KPlato
except:
    KPlato=Kross.module("KPlato")

def printNodes( node, props, schedule, types = None ):
    printNode( node, props, schedule, types )
    for i in range( node.childCount() ):
        printNodes( node.childAt( i ), props, schedule, types )

def printNode( node, props, schedule, types = None ):
    if types is None or node.type() in types:
        for prop in props:
            print "%-25s" % ( KPlato.data( node, prop[0], prop[1], schedule ) ),
        print
        print dir(node)
    else:
        print "missing: ", node.type(),dir(node)

def printSchedules(proj):
    print "%-10s %-25s %-10s" % ( "Identity", "Name", "Scheduled" )
    for i in range( proj.scheduleCount() ):
        printSchedule( proj.scheduleAt( i ) )
    print


def printSchedule( sch ):
    print "%-10s %-25s %-10s" % ( sch.id(), sch.name(), sch.isScheduled() )
    print "sch ( %s )" % ";".join(dir(sch))
    print sch.appointments()
    for i in range( sch.childCount() ):
        printSchedule( sch.childAt( i ) )

forms=Kross.module("forms")
print "forms: ",dir(forms)

for mn in dir(KPlato):
    m=getattr(KPlato,mn)
    print mn, "===>" ,m.__doc__

KPlato.document().openUrl('/home/dimon/work/python/koffice-scripts/kplato/test_project.kplato')
project=KPlato.project()

print dir(KPlato.data)

sid = -1;
# get a schedule id
if project.scheduleCount() > 0:
    sched = project.scheduleAt( 0 )
    sid = sched.id()
    print 'appointments?? : ',KPlato.data(sched,'m_appointments')

print project.id(), KPlato.data(project,'NodeName')

printSchedules(project)

nodeprops = [['NodeWBSCode', 'DisplayRole'], ['NodeName', 'DisplayRole'], ['NodeType', 'DisplayRole'], ['NodeAssignments', 'DisplayRole'], ['NodeStatus', 'EditRole'] ]
print "Print tasks and milestones in arbitrary order:"
# print the localized headers
for prop in nodeprops:
    print "%-25s" % (project.nodeHeaderData( prop ) ),

print "Nodes:"
print
printNodes( project, nodeprops, sid, [ 'Task', 'Milestone' ] )
print
