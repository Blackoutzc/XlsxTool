# -*- coding:utf-8 -*-


class Encoding:
    @staticmethod
    def to_unicode(string):
        if isinstance(string, str):
            return string.decode("utf-8")
        else:
            return string

    @staticmethod
    def to_str(string):
        if isinstance(string, str):
            return string
        else:
            return string.encode("utf-8")

    @staticmethod
    def try_to_unicode(object):
        try:
            if isinstance(object, unicode):
                return object
            return Encoding.to_unicode(Encoding.try_to_str(object))
        except Exception:
            return None

    @staticmethod
    def try_to_str(object):
        try:
            return str(object)
        except Exception:
            return None

#inherit from IUnicode class will set all value to unicode
class IUnicode(object):
    def __setattr__(self, key, value):
        print "set attribute unicode {}:{}".format(key,value)
        if isinstance(value,str):
            value = Encoding.to_unicode(value)
        super(IUnicode,self).__setattr__(key,value)


class Test(IUnicode):

    def __init__(self):
        self.age = "str"

    def __getattribute__(self, item):
        return super(Test,self).__getattribute__(item)

    def __getattr__(self, item):
        print "item"

    def get(self):
        return "get"

        #return super(Test,self).__getattr__(item)
def add(self):
    return self.age

if __name__ == '__main__':
    t = Test()

