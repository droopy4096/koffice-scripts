#!/usr/bin/python

# Get ready for Python3
from __future__ import print_function

import icalendar
import zipfile
import sys
# from xml import xpath

# from lxml import etree
# from xml.etree.ElementTree import ElementTree as etree
from xml.etree import ElementTree

class NotImplemented:
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
        raise NotImplemented

    def save(self):
        raise NotImplemented

class KPlatoFile(FormatFile):
    _resource_hash=None
    _resource_rhash=None
    _task_hash=None
    _task_rhash=None

    def __init__(self,filename):
        super(KPlatoFile,self).__init__(filename)

        self._resource_hash={}
        self._resource_rhash={}
        self._task_hash={}
        self._task_rhash={}
        self.data=None

    def load(self):
        tree=ElementTree.ElementTree()
        if self.data:
            try:
                self.data.close()
            except:
                # Brutal but simple...
                pass
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
            self._resource_hash[name]={'units':node.attrib['units'],'id':node.attrib['id']}
            self._resource_rhash[node.attrib['id']]={'units':node.attrib['units'],'name':name}
            print(name,' : ',node.attrib['id'])

        tnodes={}
        for t in tree.iter('task'):
            tnodes[t.attrib['name']]=t
            self._task_hash[t.attrib['name']]={'id':t.attrib['id']}
            self._task_rhash[t.attrib['id']]={'name':t.attrib['name']}

        print(tnodes)

        snodes={}
        for s in tree.iter('plan'):
            print(s.attrib['name'],s)
        print("we're here")


class ICalFile(FormatFile):
    def __init__(self,filename):
        super(ICalFile,self).__init__(filename)


if __name__ == "__main__":
    if len(sys.argv)<2:
        print("Usage: {0} <file.kplato> <schedule_name>".format(sys.argv[0]))
        sys.exit(1)

    kplato=KPlatoFile(sys.argv[1])
    kplato.load()
