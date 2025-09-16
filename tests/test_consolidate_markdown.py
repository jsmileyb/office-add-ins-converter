import os
import shutil
import unittest
from utils.markdown_utils import consolidate_and_clean_markdown

class TestMarkdownConsolidation(unittest.TestCase):
    def setUp(self):
        self.test_parent = 'test_parent_dir'
        self.test_output = 'test_output_dir'
        os.makedirs(self.test_parent, exist_ok=True)
        os.makedirs(self.test_output, exist_ok=True)
        child = os.path.join(self.test_parent, 'child1')
        os.makedirs(child, exist_ok=True)
        with open(os.path.join(child, 'a.md'), 'w') as f:
            f.write('Line 1\n{2}------------------------------------------------\nLine 2')
        with open(os.path.join(child, 'b.md'), 'w') as f:
            f.write('Keep this line\n{3}------------------------------------------------')

    def tearDown(self):
        shutil.rmtree(self.test_parent)
        shutil.rmtree(self.test_output)

    def test_consolidation_and_cleaning(self):
        consolidate_and_clean_markdown(self.test_parent, self.test_output)
        output_file = os.path.join(self.test_output, 'child1.md')
        self.assertTrue(os.path.exists(output_file))
        with open(output_file, 'r') as f:
            content = f.read()
            self.assertNotIn('{2}------------------------------------------------', content)
            self.assertNotIn('{3}------------------------------------------------', content)
            self.assertIn('Line 1', content)
            self.assertIn('Line 2', content)
            self.assertIn('Keep this line', content)

if __name__ == '__main__':
    unittest.main()
