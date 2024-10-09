#import azure.functions as func
import json
import inspect

def hasattr(o,n):
    return n in o.__dict__

class CONT:
    def __getattr__(self, name):
        if name in self.__dict__:
            # Default behaviour
            return self.__getattribute__(self, name)
        else:
            return None
            
#FIXME: There are two initializations of XEDB. One here and one in WappHandler.
XEDB=CONT()


#def response(mobj):
#    return func.HttpResponse(dumps(mobj))


def getIParam(req):
    req_body = req.get_body()
    return loads(req_body)



def obj2dict(pp):
    if isinstance(pp,tuple):#type(pp) is types.TupleType:
        olist=[]
        for p in pp:
            olist.append(obj2dict(p))
        return tuple(olist)
    if isinstance(pp,list):#type(pp) is types.ListType:
        olist=[]
        for p in pp:
            olist.append(obj2dict(p))
        return olist

    if isinstance(pp,dict):#type(pp) is types.DictionaryType:
        odict={}
        for k,val in pp.items():
            odict[k]=obj2dict(val)
        return odict
    if isinstance(pp,CONT): #inspect.isclass(pp): #:#type(pp) is types.InstanceType:
        odict={}
        for m in pp.__dict__:
            val=getattr(pp,m)
            odict[m]=obj2dict(val)
        return odict
            
        for m in dir(pp):
            if not m.startswith("_"): 
                val=getattr(pp,m)
                odict[m]=obj2dict(val)
        return odict
    return pp


def dict2obj(pp):
    if isinstance(pp,list):# type(pp) is types.ListType:
        olist=[]
        for p in pp:
            olist.append(dict2obj(p))
        return olist
    if isinstance(pp,dict):#type(pp) is types.DictionaryType:
        oinst=CONT()
        for (k,v) in pp.items():
            val=dict2obj(v)
            #print val
            setattr(oinst,k,val)
        return oinst
    return pp

def dumps(pp):
    return json.dumps(obj2dict(pp)) #,ensure_ascii=False)
def loads(pp):
    return dict2obj(json.loads(pp))  #,ensure_ascii=False))
