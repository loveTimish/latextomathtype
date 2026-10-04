#!/usr/bin/env python3
"""Fast process-lifecycle tests; real LibreOffice visual tests run separately."""
import os
import pathlib
import subprocess
import sys
import tempfile
import time
import unittest

HELPER = pathlib.Path(__file__).with_name('export_with_baseline.py').resolve()


class HelperTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = pathlib.Path(self.temp.name)
        self.source = self.root/'input.docx'; self.source.write_bytes(b'synthetic DOCX placeholder')
        self.jar = self.root/'uno.jar'; self.jar.write_bytes(b'fake jar for lifecycle testing only')
        self.java = self.root/'java'
        self.office = self.root/'soffice'
        self.office.write_text('#!'+sys.executable+'\nimport os,time,pathlib\npathlib.Path('+repr(str(self.root/'office.pid'))+').write_text(str(os.getpid()))\ntime.sleep(60)\n')
        self.office.chmod(0o700)

    def tearDown(self):
        self.temp.cleanup()

    def command(self, body, timeout=10):
        self.java.write_text('#!'+sys.executable+'\nimport sys,time,pathlib,os\nif "com.sun.tools.javac.Main" in sys.argv: sys.exit(0)\n'+body)
        self.java.chmod(0o700)
        return [sys.executable,str(HELPER),str(self.source),str(self.root/'output'),'--uno-jar',str(self.jar),'--java',str(self.java),'--soffice',str(self.office),'--timeout',str(timeout)]

    @staticmethod
    def success():
        return "pathlib.Path(sys.argv[sys.argv.index('LibreOfficeBaselineExporter')+3]).write_bytes(b'%PDF-native')\npathlib.Path(sys.argv[sys.argv.index('LibreOfficeBaselineExporter')+4]).write_bytes(b'%PDF-compatible')\npathlib.Path(sys.argv[sys.argv.index('LibreOfficeBaselineExporter')+5]).write_text('audit')\n"

    def assert_no_children(self):
        pidfile=self.root/'office.pid'
        if pidfile.exists():
            pid=int(pidfile.read_text())
            for _ in range(100):
                if not pathlib.Path('/proc',str(pid)).exists(): return
                time.sleep(.01)
            self.fail('private office process survived: '+str(pid))

    def test_success_unchanged_source_and_no_overwrite(self):
        command=self.command(self.success()+"print('CJK_FONT_FALLBACK synthetic mapping; portions=1')\n")
        result=subprocess.run(command,capture_output=True,timeout=15)
        self.assertEqual(0,result.returncode,result.stderr)
        self.assertIn(b'CJK_FONT_FALLBACK synthetic mapping; portions=1',result.stdout)
        self.assertEqual(b'synthetic DOCX placeholder',self.source.read_bytes())
        self.assertEqual(2,len(list((self.root/'output').glob('*.pdf'))))
        again=subprocess.run(command,capture_output=True,timeout=15)
        self.assertNotEqual(0,again.returncode)
        self.assert_no_children()

    def test_partial_failure_publishes_no_pdf(self):
        result=subprocess.run(self.command(self.success()+"sys.exit(3)\n"),capture_output=True,timeout=15)
        self.assertNotEqual(0,result.returncode)
        self.assertEqual([],list((self.root/'output').glob('*.pdf')))
        self.assert_no_children()

    def test_timeout_reaps_office_and_publishes_nothing(self):
        result=subprocess.run(self.command('time.sleep(60)\n',timeout=1),capture_output=True,timeout=10)
        self.assertEqual(124,result.returncode,result.stderr)
        self.assertEqual([],list((self.root/'output').glob('*.pdf')))
        self.assert_no_children()

    def test_signal_reaps_office_and_publishes_nothing(self):
        process=subprocess.Popen(self.command('time.sleep(60)\n'),stdout=subprocess.PIPE,stderr=subprocess.PIPE)
        deadline=time.monotonic()+5
        while not (self.root/'office.pid').exists() and time.monotonic()<deadline: time.sleep(.01)
        process.terminate(); process.communicate(timeout=10)
        self.assertEqual(124,process.returncode)
        self.assertEqual([],list((self.root/'output').glob('*.pdf')))
        self.assert_no_children()

    def test_missing_executable_is_dependency_error(self):
        command=self.command(self.success()); command[command.index('--java')+1]='/definitely/missing/java'
        result=subprocess.run(command,capture_output=True,timeout=15)
        self.assertEqual(69,result.returncode,result.stderr)
        self.assertEqual([],list((self.root/'output').glob('*.pdf')))

    def test_child_dependency_error_preserves_exit69_and_no_pdf(self):
        result=subprocess.run(self.command(self.success()+"sys.exit(69)\n"),capture_output=True,timeout=15)
        self.assertEqual(69,result.returncode,result.stderr)
        self.assertEqual([],list((self.root/'output').glob('*.pdf')))
        self.assert_no_children()


if __name__=='__main__': unittest.main()
