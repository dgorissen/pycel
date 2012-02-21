from excellib import *
from excelutil import *
from excelwrapper import ExcelComWrapper
from math import *
from networkx.classes.digraph import DiGraph
from networkx.drawing.nx_pydot import write_dot
from networkx.drawing.nx_pylab import draw, draw_circular
from networkx.readwrite.gexf import write_gexf
from tokenizer import ExcelParser, f_token, shunting_yard
import cPickle
import logging
import networkx as nx


__version__ = filter(str.isdigit, "$Revision: 2524 $")
__date__ = filter(str.isdigit, "$Date: 2011-09-06 17:05:00 +0100 (Tue, 06 Sep 2011) $")
__author__ = filter(str.isdigit, "$Author: dg2d09 $")


class Spreadsheet(object):
    def __init__(self,G,cellmap):
        super(Spreadsheet,self).__init__()
        self.G = G
        self.cellmap = cellmap
        self.params = None

    @staticmethod
    def load_from_file(fname):
        f = open(fname,'rb')
        obj = cPickle.load(f)
        #obj = load(f)
        return obj
    
    def save_to_file(self,fname):
        f = open(fname,'wb')
        cPickle.dump(self, f, protocol=2)
        f.close()

    def export_to_dot(self,fname):
        write_dot(self.G,fname)
                    
    def export_to_gexf(self,fname):
        write_gexf(self.G,fname)
    
    def plot_graph(self):
        import matplotlib.pyplot as plt

        pos=nx.spring_layout(self.G,iterations=2000)
        #pos=nx.spectral_layout(G)
        #pos = nx.random_layout(G)
        nx.draw_networkx_nodes(self.G, pos)
        nx.draw_networkx_edges(self.G, pos, arrows=True)
        nx.draw_networkx_labels(self.G, pos)
        plt.show()
    
    def set_value(self,cell,val,is_addr=True):
        if is_addr:
            cell = self.cellmap[cell]

        if cell.value != val:
            # reset the node + its dependencies
            self.reset(cell)
            # set the value
            cell.value = val
        
    def reset(self, cell):
        if cell.value is None: return
        #print "resetting", cell.address()
        cell.value = None
        map(self.reset,self.G.successors_iter(cell)) 

    def print_value_tree(self,addr,indent):
        cell = self.cellmap[addr]
        print "%s %s = %s" % (" "*indent,addr,cell.value)
        for c in self.G.predecessors_iter(cell):
            self.print_value_tree(c.address(), indent+1)

    def recalculate(self):
        for c in self.cellmap.values():
            if isinstance(c,CellRange):
                self.evaluate_range(c,is_addr=False)
            else:
                self.evaluate(c,is_addr=False)
                
    def evaluate_range(self,rng,is_addr=True):

        if is_addr:
            rng = self.cellmap[rng]

        # its important that [] gets treated ad false here
        if rng.value:
            return rng.value

        cells,nrows,ncols = rng.celladdrs,rng.nrows,rng.ncols

        if nrows == 1 or ncols == 1:
            data = [ self.evaluate(c) for c in cells ]
        else:
            data = [ [self.evaluate(c) for c in cells[i]] for i in range(len(cells)) ] 
        
        rng.value = data
        
        return data

    def evaluate(self,cell,is_addr=True):

        if is_addr:
            cell = self.cellmap[cell]
            
        # no formula, fixed value
        if not cell.formula or cell.value != None:
            #print "  returning constant or cached value for ", cell.address()
            return cell.value
        
        # recalculate formula
        # the compiled expression calls this function
        def eval_cell(address):
            return self.evaluate(address)
        
        def eval_range(rng):
            return self.evaluate_range(rng)
                
        try:
            #print "Evalling: %s, %s" % (cell.address(),cell.python_expression)
            vv = eval(cell.compiled_expression)
            #print "Cell %s evalled to %s" % (cell.address(),vv)
            if vv is None:
                print "WARNING %s is None" % (cell.address())
            cell.value = vv
        except Exception as e:
            if e.message.startswith("Problem evalling"):
                raise e
            else:
                raise Exception("Problem evalling: %s for %s, %s" % (e,cell.address(),cell.python_expression)) 
        
        return cell.value

class ASTNode(object):
    """A generic node in the AST"""
    
    def __init__(self,token):
        super(ASTNode,self).__init__()
        self.token = token
    def __str__(self):
        return self.token.tvalue
    def __getattr__(self,name):
        return getattr(self.token,name)

    def children(self,ast):
        args = ast.predecessors(self)
        args = sorted(args,key=lambda x: ast.node[x]['pos'])
        #args.reverse()
        return args

    def parent(self,ast):
        args = ast.successors(self)
        return args[0] if args else None
    
    def emit(self,ast,context=None):
        """Emit code"""
        self.token.tvalue
    
class OperatorNode(ASTNode):
    def __init__(self,*args):
        super(OperatorNode,self).__init__(*args)
        
        # convert the operator to python equivalents
        self.opmap = {
                 "^":"**",
                 "=":"==",
                 "&":"+",
                 "":"+" #union
                 }

    def emit(self,ast,context=None):
        xop = self.tvalue
        
        # Get the arguments
        args = self.children(ast)
        
        op = self.opmap.get(xop,xop)
        
        if self.ttype == "operator-prefix":
            return "-" + args[0].emit(ast,context=context)

        parent = self.parent(ast)
        # dont render the ^{1,2,..} part in a linest formula
        #TODO: bit of a hack
        if op == "**":
            if parent and parent.tvalue.lower() == "linest": 
                return args[0].emit(ast,context=context)

        #TODO silly hack to work around the fact that None < 0 is True (happens on blank cells)
        if op == "<" or op == "<=":
            aa = args[0].emit(ast,context=context)
            ss = "(" + aa + " if " + aa + " is not None else float('inf'))" + op + args[1].emit(ast,context=context)
        elif op == ">" or op == ">=":
            aa = args[1].emit(ast,context=context)
            ss =  args[0].emit(ast,context=context) + op + "(" + aa + " if " + aa + " is not None else float('inf'))"
        else:
            ss = args[0].emit(ast,context=context) + op + args[1].emit(ast,context=context)
        
        #avoid needless parentheses
        if parent and not isinstance(parent,FunctionNode):
            ss = "("+ ss + ")" 
        
        return ss

class OperandNode(ASTNode):
    def __init__(self,*args):
        super(OperandNode,self).__init__(*args)
    def emit(self,ast,context=None):
        t = self.tsubtype
        
        if t == "logical":
            return str(self.tvalue.lower() == "true")
        elif t == "text" or t == "error":
            #if the string contains quotes, escape them
            val = self.tvalue.replace('"','\\"')
            return '"' + val + '"'
        else:
            return str(self.tvalue)

class RangeNode(OperandNode):
    """Represents a spreadsheet cell or range, e.g., A5 or B3:C20"""
    def __init__(self,*args):
        super(RangeNode,self).__init__(*args)
    
    def get_cells(self):
        return resolve_range(self.tvalue)[0]
    
    def emit(self,ast,context=None):
        # resolve the range into cells
        rng = self.tvalue.replace('$','')
        sheet = context.curcell.sheet + "!" if context else ""
        if is_range(rng):
            sh,start,end = split_range(rng)
            if sh:
                str = 'eval_range("' + rng + '")'
            else:
                str = 'eval_range("' + sheet + rng + '")'
        else:
            sh,col,row = split_address(rng)
            if sh:
                str = 'eval_cell("' + rng + '")'
            else:
                str = 'eval_cell("' + sheet + rng + '")'
                
        return str
    
class FunctionNode(ASTNode):
    """AST node representing a function call"""
    def __init__(self,*args):
        super(FunctionNode,self).__init__(*args)
        self.numargs = 0

        # map excel functions onto their python equivalents, taking care to avoid name clashes
        self.funmap = {
                  "ln":"xlog",
                  "min":"xmin",
                  "min":"xmin",
                  "max":"xmax",
                  "sum":"xsum",
                  "gammaln":"lgamma"
                  }
        
    def emit(self,ast,context=None):
        xfun = self.tvalue.lower()
        str = ''

        fun = self.funmap.get(xfun,xfun)        
        
        # Get the arguments
        args = self.children(ast)
        
        if fun == "atan2":
            # swap arguments
            str = "atan2(%s,%s)" % (args[1].emit(ast,context=context),args[0].emit(ast,context=context))
        elif fun == "pi":
            # constant, no parens
            str = "pi"
        elif fun == "if":
            # inline the if
            if len(args) == 2:
                str = "%s if %s else 0" %(args[1].emit(ast,context=context),args[0].emit(ast,context=context))
            elif len(args) == 3:
                str = "(%s if %s else %s)" % (args[1].emit(ast,context=context),args[0].emit(ast,context=context),args[2].emit(ast,context=context))
            else:
                raise Exception("if with %s arguments not supported" % len(args))

        elif fun == "array":
            str += '['
            if len(args) == 1:
                # only one row
                str += args[0].emit(ast,context=context)
            else:
                # multiple rows
                str += ",".join(['[' + n.emit(ast,context=context) + ']' for n in args])
                     
            str += ']'
        elif fun == "arrayrow":
            #simply create a list
            str += ",".join([n.emit(ast,context=context) for n in args])
        elif fun == "linest" or fun == "linestmario":
            
            str = fun + "(" + ",".join([n.emit(ast,context=context) for n in args])

            if not context:
                degree,coef = -1,-1
            else:
                #linests are often used as part of an array formula spanning multiple cells,
                #one cell for each coefficient.  We have to figure out where we currently are
                #in that range
                degree,coef = get_linest_degree(context.excel,context.curcell)
                
            # if we are the only linest (degree is one) and linest is nested -> return vector
            # else return the coef.
            if degree == 1 and self.parent(ast):
                if fun == "linest":
                    str += ",degree=%s)" % degree
                else:
                    str += ")"
            else:
                if fun == "linest":
                    str += ",degree=%s)[%s]" % (degree,coef-1)
                else:
                    str += ")[%s]" % (coef-1)

        elif fun == "and":
            str = "all([" + ",".join([n.emit(ast,context=context) for n in args]) + "])"
        elif fun == "or":
            str = "any([" + ",".join([n.emit(ast,context=context) for n in args]) + "])"
        else:
            str = fun + "(" + ",".join([n.emit(ast,context=context) for n in args]) + ")"

        return str

def create_node(t):
    """Simple factory function"""
    if t.ttype == "operand":
        if t.tsubtype == "range":
            return RangeNode(t)
        else:
            return OperandNode(t)
    elif t.ttype == "function":
        return FunctionNode(t)
    elif t.ttype.startswith("operator"):
        return OperatorNode(t)
    else:
        return ASTNode(t)

class Operator:
    """Small wrapper class to manage operators during shunting yard"""
    def __init__(self,value,precedence,associativity):
        self.value = value
        self.precedence = precedence
        self.associativity = associativity

def shunting_yard(expression):
    """
    Tokenize an excel formula expression into reverse polish notation
    
    Core algorithm taken from wikipedia with varargs extensions from
    http://www.kallisti.net.nz/blog/2008/02/extension-to-the-shunting-yard-algorithm-to-allow-variable-numbers-of-arguments-to-functions/
    """
    #remove leading =
    if expression.startswith('='):
        expression = expression[1:]
        
    p = ExcelParser();
    p.parse(expression)

    # insert tokens for '(' and ')', to make things clearer below
    tokens = []
    for t in p.tokens.items:
        if t.ttype == "function" and t.tsubtype == "start":
            t.tsubtype = ""
            tokens.append(t)
            tokens.append(f_token('(','arglist','start'))
        elif t.ttype == "function" and t.tsubtype == "stop":
            tokens.append(f_token(')','arglist','stop'))
        elif t.ttype == "subexpression" and t.tsubtype == "start":
            t.tvalue = '('
            tokens.append(t)
        elif t.ttype == "subexpression" and t.tsubtype == "stop":
            t.tvalue = ')'
            tokens.append(t)
        else:
            tokens.append(t)

    #print "tokens: ", "|".join([x.tvalue for x in tokens])

    #http://office.microsoft.com/en-us/excel-help/calculation-operators-and-precedence-HP010078886.aspx
    operators = {}
    operators[':'] = Operator(':',8,'left')
    operators[''] = Operator(' ',8,'left')
    operators[','] = Operator(',',8,'left')
    operators['u-'] = Operator('u-',7,'left') #unary negation
    operators['%'] = Operator('%',6,'left')
    operators['^'] = Operator('^',5,'left')
    operators['*'] = Operator('*',4,'left')
    operators['/'] = Operator('/',4,'left')
    operators['+'] = Operator('+',3,'left')
    operators['-'] = Operator('-',3,'left')
    operators['&'] = Operator('&',2,'left')
    operators['='] = Operator('=',1,'left')
    operators['<'] = Operator('<',1,'left')
    operators['>'] = Operator('>',1,'left')
    operators['<='] = Operator('<=',1,'left')
    operators['>='] = Operator('>=',1,'left')
    operators['<>'] = Operator('<>',1,'left')
            
    output = collections.deque()
    stack = []
    were_values = []
    arg_count = []
    
    for t in tokens:
        if t.ttype == "operand":

            output.append(create_node(t))
            if were_values:
                were_values.pop()
                were_values.append(True)
                
        elif t.ttype == "function":

            stack.append(t)
            arg_count.append(0)
            if were_values:
                were_values.pop()
                were_values.append(True)
            were_values.append(False)
            
        elif t.ttype == "argument":
            
            while stack and (stack[-1].tsubtype != "start"):
                output.append(create_node(stack.pop()))   
            
            if were_values.pop(): arg_count[-1] += 1
            were_values.append(False)
            
            if not len(stack):
                raise Exception("Mismatched or misplaced parentheses")
        
        elif t.ttype.startswith('operator'):

            if t.ttype.endswith('-prefix') and t.tvalue =="-":
                o1 = operators['u-']
            else:
                o1 = operators[t.tvalue]

            while stack and stack[-1].ttype.startswith('operator'):
                
                if stack[-1].ttype.endswith('-prefix') and stack[-1].tvalue =="-":
                    o2 = operators['u-']
                else:
                    o2 = operators[stack[-1].tvalue]
                
                if ( (o1.associativity == "left" and o1.precedence <= o2.precedence)
                        or
                      (o1.associativity == "right" and o1.precedence < o2.precedence) ):
                    
                    output.append(create_node(stack.pop()))
                else:
                    break
                
            stack.append(t)
        
        elif t.tsubtype == "start":
            stack.append(t)
            
        elif t.tsubtype == "stop":
            
            while stack and stack[-1].tsubtype != "start":
                output.append(create_node(stack.pop()))
            
            if not stack:
                raise Exception("Mismatched or misplaced parentheses")
            
            stack.pop()

            if stack and stack[-1].ttype == "function":
                f = create_node(stack.pop())
                a = arg_count.pop()
                w = were_values.pop()
                if w: a += 1
                f.num_args = a
                #print f, "has ",a," args"
                output.append(f)

    while stack:
        if stack[-1].tsubtype == "start" or stack[-1].tsubtype == "stop":
            raise Exception("Mismatched or misplaced parentheses")
        
        output.append(create_node(stack.pop()))

    #print "Stack is: ", "|".join(stack)
    #print "Ouput is: ", "|".join([x.tvalue for x in output])
    
    # convert to list
    result = [x for x in output]
    return result
   
def build_ast(expression):
    """build an AST from an Excel formula expression in reverse polish notation"""
    
    #use a directed graph to store the tree
    G = DiGraph()
    
    stack = []
    
    for n in expression:
        # Since the graph does not maintain the order of adding nodes/edges
        # add an extra attribute 'pos' so we can always sort to the correct order
        if isinstance(n,OperatorNode):
            if n.ttype == "operator-infix":
                arg2 = stack.pop()
                arg1 = stack.pop()
                G.add_node(arg1,{'pos':1})
                G.add_node(arg2,{'pos':2})
                G.add_edge(arg1, n)
                G.add_edge(arg2, n)
            else:
                arg1 = stack.pop()
                G.add_node(arg1,{'pos':1})
                G.add_edge(arg1, n)
                
        elif isinstance(n,FunctionNode):
            args = [stack.pop() for _ in range(n.num_args)]
            args.reverse()
            for i,a in enumerate(args):
                G.add_node(a,{'pos':i})
                G.add_edge(a,n)
            #for i in range(n.num_args):
            #    G.add_edge(stack.pop(),n)
        else:
            G.add_node(n,{'pos':0})

        stack.append(n)
        
    return G,stack.pop()

class Context(object):
    """A small context object that nodes in the AST can use to emit code"""
    def __init__(self,curcell,excel):
        # the current cell for which we are generating code
        self.curcell = curcell
        # a handle to an excel instance
        self.excel = excel

class ExcelCompiler(object):
    """Class responsible for taking an Excel spreadsheet and compiling it to a Spreadsheet instance
       that can be serialized to disk, and executed independently of excel.
       
       Must be run on Windows as it requires a COM link to an Excel instance.
       """
       
    def __init__(self, filename=None, excel=None, *args,**kwargs):

        super(ExcelCompiler,self).__init__()
        self.filename = filename
        
        if excel:
            # if we are running as an excel addin, this gets passed to us
            self.excel = excel
        else:
            # TODO: use a proper interface so we can (eventually) support loading from file (much faster)  Still need to find a good lib though.
            self.excel = ExcelComWrapper(filename=filename)
            self.excel.connect()
            
        self.log = logging.getLogger("decode.{0}".format(self.__class__.__name__))
        
    def cell2code(self,cell):
        """Generate python code for the given cell"""
        if cell.formula:
            e = shunting_yard(cell.formula or str(cell.value))
            ast,root = build_ast(e)
            code = root.emit(ast,context=Context(cell,self.excel))
        else:
            ast = None
            code = str('"' + cell.value + '"' if isinstance(cell.value,unicode) else cell.value)
            
        return code,ast

    def add_node_to_graph(self,G, n):
        G.add_node(n)
        G.node[n]['sheet'] = n.sheet
        
        if isinstance(n,Cell):
            G.node[n]['label'] = n.col + str(n.row)
        else:
            #strip the sheet
            G.node[n]['label'] = n.address()[n.address().find('!')+1:]
            
    def gen_graph(self, seed, sheet=None):
        """Given a starting point (e.g., A6, or A3:B7) on a particular sheet, generate
           a Spreadsheet instance that captures the logic and control flow of the equations."""
        
        # starting points
        cursheet = sheet if sheet else self.excel.get_active_sheet()
        self.excel.set_sheet(cursheet)
        
        seeds,nr,nc = Cell.make_cells(self.excel, seed, sheet=cursheet)
        seeds = list(flatten(seeds))
        
        print "Seed %s expanded into %s cells" % (seed,len(seeds))
        
        # only keep seeds with formulas or numbers
        seeds = [s for s in seeds if s.formula or isinstance(s.value,(int,float))]

        print "%s filtered seeds " % len(seeds)
        
        # cells to analyze: only formulas
        todo = [s for s in seeds if s.formula]

        print "%s cells on the todo list" % len(todo)

        # map of all cells
        cellmap = dict([(x.address(),x) for x in seeds])
        
        # directed graph
        G = nx.DiGraph()

        # match the info in cellmap
        for c in cellmap.itervalues(): self.add_node_to_graph(G, c)

        while todo:
            c1 = todo.pop()
            
            print "Handling ", c1.address()
            
            # set the current sheet so relative addresses resolve properly
            if c1.sheet != cursheet:
                cursheet = c1.sheet
                self.excel.set_sheet(cursheet)
            
            # parse the formula into code
            pystr,ast = self.cell2code(c1)

            # set the code & compile it (will flag problems sooner rather than later)
            c1.python_expression = pystr
            c1.compile()                
            
            # get all the cells/ranges this formula refers to
            deps = [x.tvalue.replace('$','') for x in ast.nodes() if isinstance(x,RangeNode)]
            
            # remove dupes
            deps = uniqueify(deps)
            
            for dep in deps:
                
                # if the dependency is a multi-cell range, create a range object
                if is_range(dep):
                    # this will make sure we always have an absolute address
                    rng = CellRange(dep,sheet=cursheet)
                    
                    if rng.address() in cellmap:
                        # already dealt with this range
                        # add an edge from the range to the parent
                        G.add_edge(cellmap[rng.address()],cellmap[c1.address()])
                        continue
                    else:
                        # turn into cell objects
                        cells,nrows,ncols = Cell.make_cells(self.excel,dep,sheet=cursheet)

                        # get the values so we can set the range value
                        if nrows == 1 or ncols == 1:
                            rng.value = [c.value for c in cells]
                        else:
                            rng.value = [ [c.value for c in cells[i]] for i in range(len(cells)) ] 

                        # save the range
                        cellmap[rng.address()] = rng
                        # add an edge from the range to the parent
                        self.add_node_to_graph(G, rng)
                        G.add_edge(rng,cellmap[c1.address()])
                        # cells in the range should point to the range as their parent
                        target = rng
                else:
                    # not a range, create the cell object
                    cells = [Cell.resolve_cell(self.excel, dep, sheet=cursheet)]
                    target = cellmap[c1.address()]

                # process each cell                    
                for c2 in flatten(cells):
                    # if we havent treated this cell allready
                    if c2.address() not in cellmap:
                        if c2.formula:
                            # cell with a formula, needs to be added to the todo list
                            todo.append(c2)
                            #print "appended ", c2.address()
                        else:
                            # constant cell, no need for further processing, just remember to set the code
                            pystr,ast = self.cell2code(c2)
                            c2.python_expression = pystr
                            c2.compile()     
                            #print "skipped ", c2.address()
                        
                        # save in the cellmap
                        cellmap[c2.address()] = c2
                        # add to the graph
                        self.add_node_to_graph(G, c2)
                        
                    # add an edge from the cell to the parent (range or cell)
                    G.add_edge(cellmap[c2.address()],target)
            
        print "Graph construction done, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap))

        sp = Spreadsheet(G,cellmap)
        
        return sp

if __name__ == '__main__':
    
    # some test formulas
    inputs = [
              '=SUM((A:A 1:1))',
              '=A1',
              '=atan2(A1,B1)',
              '=5*log(sin()+2)',
              '=5*log(sin(3,7,9)+2)',
              '=3 + 4 * 2 / ( 1 - 5 ) ^ 2 ^ 3',
              '=1+3+5',
              '=3 * 4 + 5',
              '=50',
              '=1+1',
              '=$A1',
              '=$B$2',
              '=SUM(B5:B15)',
              '=SUM(B5:B15,D5:D15)',
              '=SUM(B5:B15 A7:D7)',
              '=SUM(sheet1!$A$1:$B$2)',
              '=[data.xls]sheet1!$A$1',
              '=SUM((A:A,1:1))',
              '=SUM((A:A A1:B1))',
              '=SUM(D9:D11,E9:E11,F9:F11)',
              '=SUM((D9:D11,(E9:E11,F9:F11)))',
              '=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))',
              '={SUM(B2:D2*B3:D3)}',
              '=SUM(123 + SUM(456) + (45<6))+456+789',
              '=AVG(((((123 + 4 + AVG(A1:A2))))))',
              
              # E. W. Bachtal's test formulae
              '=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &   "  more ""test"" text"',
              #'=+ AName- (-+-+-2^6) = {"A","B"} + @SUM(R1C1) + (@ERROR.TYPE(#VALUE!) = 2)',
              '=IF(R13C3>DATE(2002,1,6),0,IF(ISERROR(R[41]C[2]),0,IF(R13C3>=R[41]C[2],0, IF(AND(R[23]C[11]>=55,R[24]C[11]>=20),R53C3,0))))',
              '=IF(R[39]C[11]>65,R[25]C[42],ROUND((R[11]C[11]*IF(OR(AND(R[39]C[11]>=55, ' + 
                  'R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")),R[44]C[11],R[43]C[11]))+(R[14]C[11] ' +
                  '*IF(OR(AND(R[39]C[11]>=55,R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")), ' +
                  'R[45]C[11],R[43]C[11])),0))',
              '=(propellor_charts!B22*(propellor_charts!E21+propellor_charts!D21*(engine_data!O16*D70+engine_data!P16)+propellor_charts!C21*(engine_data!O16*D70+engine_data!P16)^2+propellor_charts!B21*(engine_data!O16*D70+engine_data!P16)^3)^2)^(1/3)*(1*D70/5.33E-18)^(2/3)*0.0000000001*28.3495231*9.81/1000',
              '=(3600/1000)*E40*(E8/E39)*(E15/E19)*LN(E54/(E54-E48))',
              '=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))',
              '=LINEST(X5:X32,W5:W32^{1,2,3})',
              '=IF(configurations!$G$22=3,sizing!$C$303,M14)',
              '=0.000001042*E226^3-0.00004777*E226^2+0.0007646*E226-0.00075',
              '=LINEST(G2:G17,E2:E17,FALSE)',
              '=IF(AI119="","",E119)',
              '=LINEST(B32:(INDEX(B32:B119,MATCH(0,B32:B119,-1),1)),(F32:(INDEX(B32:F119,MATCH(0,B32:B119,-1),5)))^{1,2,3,4})',
              ]

    for i in inputs:
        print "**************************************************"
        print "Formula: ", i

        e = shunting_yard(i);
        print "RPN: ",  "|".join([str(x) for x in e])
        
        G,root = build_ast(e)
        
        print "Python code: ", root.emit(G,context=None)
        print "**************************************************"
