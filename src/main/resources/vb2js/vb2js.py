"""
VB Script to JavaScript convertor
"""

__author__ = 'nagakishore_byr@infosys.com'

try:
    import simplejson
except ImportError:
    import json as simplejson

import string
import os

ThisDir = os.path.abspath(os.path.dirname(__file__))

MapFileName = os.path.join(ThisDir, 'map.json')
VbsFileName = os.path.join(ThisDir, 'test.vbs.txt')
JsFileName = os.path.join(ThisDir, 'test.js')


class WordMap(dict):
    """
    A thin wrapper around built-in Dictionary.
    Loads Map file as Dictionary

    args: String, absolute path of Map (JSON) file
    """
    def __init__(self, mapfile):
        with open(mapfile, 'rb') as fp:
            data = simplejson.load(fp)
        super(WordMap, self).__init__(data)

    def get(self, name, rule=None, default=None):
        """
        args
            name: String, key in map file 
            rule: String, key in map[name]
            default: Any object, returned when value is not found 
                     for name, rule in map file
        returns
            String if self[name][rule] or self[name] is present
            otherwise, default is returned
        """
        try:
            data = self[name][rule] if rule else self[name]
        except KeyError:
            data = default
        return data

class Line(str):
    """
    A line in VB Script 

    Usage
        >>> print Line(vb_script_line).parse()
    """
    # Load the Vb2Js JSON map
    wordmap = WordMap(MapFileName)

    def __init__(self, line):
        super(Line, self).__init__(line)
        self.parsedWords = []
        
    def split(self):
        """
        Splits the line into words delimited by any of the characters
            whitespace or '(),.:=

        returns
            Array of Strings along with delimiters!
        """
        words = []
        word = ''
        for ch in self:
            if ch in string.whitespace + '()\',.:=':
                words.append(word)
                words.append(ch)
                word = ''
            else:
                word += ch
        words.append(word)
        return words

    def parse(self):
        """
        Parses the line (VB Script Line) to JavaaScript line

        TODO: Some parts are not neatly designed. For instance, identifying
        a line as function
        """
        words = self.split()
        lineBefore = ''
        lineAfter = ''
        for w in words:
            # search for any rules defined in the wordmap for word w
            # If the rule, starting with `on` is present in Line string
            # get the rule key. There may be several rules <order of precidence>
            # is yet to be decided.
            rule = None
            for key in Line.wordmap.get(w, default={}).keys():
                if key.startswith('on') and key.split('on')[1] in words:
                    rule = key
            
            # Get the data of a word for a identified rule (default None)
            # Each rule defined should have a attribute `value` in the wordmap
            data = Line.wordmap.get(w, rule=rule)

            # If the word in the line is not defined in the map, just use the word"onExit": "break", 
            # Of course, not all words have definitions
            if data is None:
                self.parsedWords.append(w)
                continue

            jsword = data.get('value', w)

            # Add before & after key values
            # Is there a more elegant place to put these whitespaces?
            jsword = data.get('before', '') + jsword + data.get('after', '')

            # If lineBefore & lineAfter key value exist
            lineBefore = data.get('lineBefore', lineBefore)
            lineAfter = data.get('lineAfter', lineAfter)
            
            # store the jsword
            self.parsedWords.append(jsword)
        
        # Identify if the line is function call or not
        # This is quite interesting syntax of VB Script
        lineCallable = True
        for key in self.wordmap.keys():
            if key in [':', '=']:
                continue
            if key in words:
                lineCallable = False
                break

        if lineCallable:
            # Put a opening bracket after the first word
            # and closing bracket after the last word.
            openAt = 0
            closeAt = 0
            for i, word in enumerate(self.parsedWords):
                if word in string.whitespace + '()\',.:=':
                    continue
                openAt = i if openAt == 0 else openAt
                closeAt = i

            # Insert the brackets if its not a new line
            if openAt > 0: 
                self.parsedWords.insert(openAt+1, '(')
                self.parsedWords.insert(closeAt+2, ')')
        
        return lineBefore + "".join(self.parsedWords).rstrip() + lineAfter + '\n'


def Main():
    vbsFile = open(VbsFileName, "rb")
    jsFile = open(JsFileName, "wb")

    for vbLine in vbsFile:
       jsFile.write(Line(vbLine).parse())
    
    vbsFile.close()
    jsFile.close()

if __name__ == '__main__':
    Main()
