#!/usr/bin/env python3

import unittest
from mastermind_puzzlegenerator import Sheet


class SheetTestCase(unittest.TestCase):
    def setUp(self):
        class A(object):
            debug = False
            pass

        class S(Sheet):
            def __init__(self,
                         correct=None,
                         clue_lines=None,
                         clue_answers=None,
                         easy=False):
                self.args = A()
                self.correct = correct
                self.clue_lines = clue_lines
                self.clue_answers = clue_answers
                self.easy = easy

        self.sheet = S()

    def testAnswerCorrectLocationBlacks(self):
        self.sheet.args.columns = 4
        self.sheet.correct = [1, 2, 3, 4]
        self.assertTupleEqual(self.sheet.answer([1, 2, 3, 4]), (4, 0,))
        self.assertTupleEqual(self.sheet.answer([1, 2, 3, 5]), (3, 0,))
        self.assertTupleEqual(self.sheet.answer([1, 2, 5, 5]), (2, 0,))
        self.assertTupleEqual(self.sheet.answer([1, 5, 5, 5]), (1, 0,))
        self.assertTupleEqual(self.sheet.answer([5, 6, 7, 8]), (0, 0,))
        self.assertTupleEqual(self.sheet.answer([5, 6, 7, 4]), (1, 0,))

    def testAnswerIncorrectLocationWhites(self):
        self.sheet.args.columns = 4
        self.sheet.correct = [1, 2, 3, 4]
        self.assertTupleEqual(self.sheet.answer([4, 1, 2, 3]), (0, 4,))
        self.assertTupleEqual(self.sheet.answer([5, 1, 2, 3]), (0, 3,))
        self.assertTupleEqual(self.sheet.answer([5, 1, 2, 5]), (0, 2,))
        self.assertTupleEqual(self.sheet.answer([5, 5, 5, 1]), (0, 1,))

    def assertBothWays(self, first, second, result):
        self.sheet.correct = first
        self.assertTupleEqual(self.sheet.answer(second), result)
        self.sheet.correct = second
        self.assertTupleEqual(self.sheet.answer(first), result)

    def testAnswerMultiValuesInCorrect(self):
        self.sheet.args.columns = 4
        first = [1, 1, 1, 1]
        self.assertBothWays(first, [1, 1, 1, 1], (4, 0,))
        self.assertBothWays(first, [1, 1, 1, 2], (3, 0,))
        self.assertBothWays(first, [1, 1, 2, 3], (2, 0,))
        self.assertBothWays(first, [2, 3, 4, 1], (1, 0,))

        first = [1, 1, 2, 2]
        self.assertBothWays(first, [1, 1, 1, 1], (2, 0,))
        self.assertBothWays(first, [1, 1, 1, 2], (3, 0,))
        self.assertBothWays(first, [1, 1, 2, 3], (3, 0,))
        self.assertBothWays(first, [2, 3, 4, 1], (0, 2,))

    def testSimplerAnswers(self):
        self.sheet.args.columns = 2
        first = [1, 2]
        self.assertBothWays(first, [1, 2], (2, 0,))

    def testCombinations(self):
        self.sheet.args.columns = 2
        self.sheet.args.colors = 2
        self.sheet.args.debug = True
        self.sheet.correct = [1, 2]
        self.assertEqual(self.sheet.combinations([]), 4)
        self.assertEqual(self.sheet.combinations([[1, 3]]), 2)


if __name__ == '__main__':
    unittest.main()
