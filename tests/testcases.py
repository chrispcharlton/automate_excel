from collections import namedtuple
import pandas as pd


PaddedTupleTestCase = namedtuple('PaddedTupleTestCase', ('values','x', 'y', 'expected'))
padded_tuple_tests = [
                PaddedTupleTestCase('value', 1, 1, (('value',),)),
                PaddedTupleTestCase('value', 1, 2, (('value', None),)),
                PaddedTupleTestCase('value', 2, 1, (('value',), (None,))),
                PaddedTupleTestCase(('a', 'b'), 1, 2, (('a','b'),)),
                PaddedTupleTestCase((('a',), ('b',)), 2, 1, (('a',), ('b',),)),
                PaddedTupleTestCase((('a','b'), ('c',)), 2, 2, (('a','b'), ('c', None),)),
                PaddedTupleTestCase((('a', 'b'), ('c',), ('d', 'e')), 3, 2, (('a', 'b'), ('c', None), ('d', 'e'))),
                PaddedTupleTestCase((('a', 'b'), 'c', ('d', 'e')), 3, 2, (('a', 'b'), ('c', None), ('d', 'e'))),
                PaddedTupleTestCase((('a', 'b'), ('c',)), 3, 2, (('a', 'b'), ('c', None), (None, None))),

]

RangeTestCase = namedtuple('RangeTestCase', ('range', 'values', 'expected_values'))
range_tests = [
                RangeTestCase('A1', 'Test', 'Test'),
                RangeTestCase('A1', 1, 1),
                RangeTestCase('A1', True, True),
                RangeTestCase('A1', pd.DataFrame([[1]]), 1),
                RangeTestCase('A1:C1', [1,2,3], ((1,2,3),)),
                RangeTestCase('A1:B2', [[1,2],[3,4]], ((1,2),(3,4))),
                RangeTestCase('A1:C2', [[1,2],1], ((1,2,None),(1, None, None))),
                RangeTestCase('A1:B3', [[1, 2], 1], ((1, 2), (1, None), (None, None))),
                RangeTestCase('A1:B2', pd.DataFrame([[1,2], [3,4]]), ((1,2),(3,4))),
                RangeTestCase('A1:B2', 1, ((1,None),(None,None))),
                ]

range_tests_fail = [
                RangeTestCase('A1', [1,2], None),
                RangeTestCase('A1:B2', [[1, 2], [3, 4], [5, 6]], None),
                RangeTestCase('A1:B2', [[1, 2, 3], [4, 5, 6]], None),
                RangeTestCase('A1:B2', pd.DataFrame([[1, 2], [3, 4], [5, 6]]), None),
                RangeTestCase('A1:B2', pd.DataFrame([[1, 2, 3], [4, 5, 6]]), None),
                ]