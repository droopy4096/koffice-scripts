#!/usr/bin/python

# Get ready for Python3
from __future__ import print_function

import zipfile
import sys
import re
# from xml import xpath

# from lxml import etree
# from xml.etree.ElementTree import ElementTree as etree
from xml.etree import ElementTree

from datetime import datetime,tzinfo,timedelta
from icalendar import Calendar, Event,vDatetime
from dateutil import parser as du_parser
# import icalendar


import time as _time

ZERO = timedelta(0)
HOUR = timedelta(hours=1)

STDOFFSET = timedelta(seconds = -_time.timezone)
if _time.daylight:
    DSTOFFSET = timedelta(seconds = -_time.altzone)
else:
    DSTOFFSET = STDOFFSET

DSTDIFF = DSTOFFSET - STDOFFSET

class LocalTimezone(tzinfo):
    """LocalTimezone object convenient for datetime manipulations"""
    def utcoffset(self, dt):
        if self._isdst(dt):
            return DSTOFFSET
        else:
            return STDOFFSET

    def dst(self, dt):
        if self._isdst(dt):
            return DSTDIFF
        else:
            return ZERO

    def tzname(self, dt):
        return _time.tzname[self._isdst(dt)]

    def _isdst(self, dt):
        tt = (dt.year, dt.month, dt.day,
              dt.hour, dt.minute, dt.second,
              dt.weekday(), 0, 0)
        stamp = _time.mktime(tt)
        tt = _time.localtime(stamp)
        return tt.tm_isdst > 0

def extract_datetime(datetime_str):
    """Extract datetime object from KPlato's datetime string
    NOTE: it doesn't work (discards) with timezones!!!"""
    res1=du_parser.parse(datetime_str)
    res2=datetime(res1.year,res1.month,res1.day,res1.hour,res1.minute,res1.second,res1.microsecond,tzinfo=LocalTimezone())
    print(res2)
    return res1
    start_str=datetime_str
    start_date=start_str[0:10]
    start_time_full=start_str[11:]
    start_time_time=start_time_full[:-6]
    if(len(start_time_time)<9):
        # print(start_time_time)
        start_time_time=start_time_time+".0"
    start_time_tz=start_time_full[-6:]
    start_time_tz=start_time_tz[:3]+start_time_tz[4:]
    # print(start_str,start_date,start_time_time,start_time_tz,sep=' | ')
    # t_start=datetime.strptime(start_str,'%Y-%m-%dT%H:%m:%S.%f%z')
    t_start=datetime.strptime(start_date+" "+start_time_time,'%Y-%m-%d %X.%f')
    return t_start

class ErrorNotImplemented:
    pass

class ErrorWrongParams:
    pass

class FormatFile(object):
    _filename=None

    _data=None

    def __init__(self,filename):
        self.set_filename(filename)

    # covering up _filename with "filename"
    def set_filename(self,filename):
        self._filename=filename

    def get_filename(self):
        return self._filename

    filename=property(get_filename,set_filename)

    # covering up _data
    def set_data(self,data):
        self._data=data

    def get_data(self):
        return self._data

    data=property(get_data,set_data)

    def load(self):
        raise ErrorNotImplemented

    def save(self):
        raise ErrorNotImplemented

class KPlatoFile(FormatFile):
    _resource_hash=None
    _resource_rhash=None
    _task_hash=None
    _task_rhash=None
    _appointment_hash=None

    def __init__(self,filename):
        super(KPlatoFile,self).__init__(filename)

        self._resource_hash={}
        self._resource_rhash={}
        self._task_hash={}
        self._task_rhash={}
        self._appointment_hash={}
        self.data=None

    def list_resource_names(self):
        return self._resource_hash.keys()

    def list_resource_ids(self):
        return self._resource_rhash.keys()

    def get_resource(self,**kwargs):
        if kwargs.has_key('name'):
            return self._resource_hash[kwargs['name']]
        elif kwargs.has_key('id'):
            return self._resource_rhash[kwargs['id']]
        else:
            raise ErrorWrongParams

    def add_resource(self,r_name,r_id,r_units):
        if not self._resource_hash.has_key(r_name):
            self.update_resource(r_name,r_id,r_units)

    def update_resource(self,r_name,r_id,r_units):
        self._resource_hash[r_name]={'units':r_units,'id':r_id}
        self._resource_rhash[r_id]={'units':r_units,'name':r_name}
        
    def del_resource(self,**kwargs):
        if kwargs['name']:
            r_name=kwargs['name']
            r_id=self._resource_hash[r_name]['id']
        elif kwargs['id']:
            r_id=kwargs['id']
            r_name=self._resource_rhash[r_id]['name']
        else:
            raise ErrorWrongParams
        del  self._resource_hash[r_name]
        del  self._resource_rhash[r_id]

    def get_task(self,**kwargs):
        if kwargs.has_key('name'):
            return self._task_hash[kwargs['name']]
        elif kwargs.has_key('id'):
            return self._task_rhash[kwargs['id']]
        else:
            raise ErrorWrongParams

    def get_appointment(self,schedule,a_tupple):
        """returns dict of intervals: {'start':,'end':,'load':}"""
        return self._appointment_hash[schedule][a_tupple]

    def list_schedules(self):
        return self._appointment_hash.keys()

    def list_appointments(self,schedule):
        """Return list: [(task-id,resource-id),...]"""
        return self._appointment_hash[schedule].keys()

    def load(self):
        tree=ElementTree.ElementTree()
        with zipfile.ZipFile(self.filename,'r') as z:
            with z.open('maindoc.xml') as kplato_data:
                # a=kplato_data.read()
                # print(a)
                self.data=tree.parse(kplato_data)
        # print(ElementTree.tostring(self.data))

        #############
        # Resources
        rnodes={}
        ## .find() and .findall work only for immediate childen...
        for r in tree.iter('resource'):
            rnodes[r.attrib['name']]={'rnode':r}
        for g in tree.iter('resource-group'):
            rs=g.findall('resource')
            for r in rs:
                rnodes[r.attrib['name']]={'rnode':r, 'gnode':g}
        for name in rnodes.keys():
            node=rnodes[name]['rnode']
            self.add_resource(name,node.attrib['id'],node.attrib['units'])
            # self._resource_hash[name]={'units':node.attrib['units'],'id':node.attrib['id']}
            # self._resource_rhash[node.attrib['id']]={'units':node.attrib['units'],'name':name}
            # print(name,' : ',node.attrib['id'])

        tnodes={}
        for t in tree.iter('task'):
            tnodes[t.attrib['name']]=t
            self._task_hash[t.attrib['name']]={'id':t.attrib['id']}
            self._task_rhash[t.attrib['id']]={'name':t.attrib['name']}

        # print(tnodes)

        snodes={}
        anodes={}
        ahash={}
        for p in tree.iter('plan'):
            pname=p.attrib['name']
            # print(pname,p)
            anodes[pname]={}
            ahash[pname]={}
            for a in p.iter('appointment'):
                task_id=a.attrib['task-id']
                resource_id=a.attrib['resource-id']
                a_id=(task_id,resource_id)
                anodes[pname][a_id]=[]
                ahash[pname][a_id]=[]
                for i in a.findall('interval'):
                    anodes[pname][a_id].append(i)
                    ahash[pname][a_id].append({'start':i.attrib['start'],'end':i.attrib['end'],'load':i.attrib['load']})
        self._appointment_hash=ahash

class ICalFile(FormatFile):
    _cal=None
    
    def __init__(self,name):
        filename=name+'.ics'
        super(ICalFile,self).__init__(filename)
        self._cal=Calendar()
        self._cal.add('prodid', '-//KPlato Export//athabascau.ca//')
        self._cal.add('version', '2.0')
    
    def add_appointment(self,uid,dtstart,dtend,dtstamp,summary,description):
        event=Event()
        event.add('summary',summary)
        event.add('description',description)
        event['uid']=uid

        # event.add('dtstart',dtstart)
        event['dtstart']=vDatetime(dtstart).ical()
        # event.add('dtend',dtend)
        event['dtend']=vDatetime(dtend).ical()
        # event.add('dtstamp',dtstamp)
        event['dtstamp']=vDatetime(dtstamp).ical()
        self._cal.add_component(event)
    
    def save(self):
        f=open(self.filename,'w')
        f.write(self._cal.as_string())
        f.close()

class KPlato2Ical(object):
    def __init__(self):
        pass

    def compile_ical(self,kplato,schedule):
        # print(kplato.list_schedules())
        # print(kplato.list_appointments(schedule))
        appts=kplato.list_appointments(schedule)
        r_cal={}
        uid=0
        for rid in kplato.list_resource_ids():
            r=kplato.get_resource(id=rid)
            # print('Resource: ',r)
            r_t=[]
            r_cal[rid]={}
            for a in appts:
                a_tid=a[0]
                a_rid=a[1]
                if a_rid == rid:
                    # this resource's task
                    t=kplato.get_task(id=a_tid)
                    r_t.append(t)
                    blocks=kplato.get_appointment(schedule,a)
                    for b in blocks:
                        # print('{0} : {1} : {2} - {3} : {4}'.format(r['name'],t['name'],b['start'],b['end'],b['load']))
                        timeslot=(b['start'],b['end'])
                        if not r_cal[rid].has_key(timeslot):
                            r_cal[rid][timeslot]=[]
                        r_cal[rid][timeslot].append({'load':b['load'],'resource-id':rid,'resource-name':r['name'],\
                                                     'task-id':a_tid,'task-name':t['name']})

            # tz_re=re.compile(r'(..)\:(..)$')
            
            icf=ICalFile(r['name'])
            for ts in r_cal[rid].keys():
                # icf.add_appointment(....)
                # print(r_cal[rid][ts])
                blocks=r_cal[rid][ts]
                t_start=extract_datetime(ts[0])
                t_end=extract_datetime(ts[1])
                t_stamp=datetime.now()
                ## add_appointment(self,uid,dtstart,dtend,dtstamp,summary,description):
                uid=uid+1
                summaries=["{0}% {1}".format(b['load'],b['task-name']) for b in blocks]

                icf.add_appointment(uid,t_start,t_end,t_stamp,"\n".join(summaries),'description')

            icf.save()
            # print('Tasks: ',r_t)

if __name__ == "__main__":
    if len(sys.argv)<2:
        print("Usage: {0} <file.kplato> <schedule_name>".format(sys.argv[0]))
        sys.exit(1)

    kplato=KPlatoFile(sys.argv[1])
    kplato.load()
    k2i=KPlato2Ical()
    k2i.compile_ical(kplato,sys.argv[2])
