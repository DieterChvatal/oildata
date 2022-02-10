import os
import sys
import subprocess
from os.path import join, pathsep
import config
import logging

path = os.path.dirname(config.wrapper_logfile)
if not os.path.exists(path):
    os.makedirs(path)

logging.basicConfig(filename=config.wrapper_logfile, format='%(asctime)s %(levelname)s, %(message)s',level=logging.DEBUG)
#logging.basicConfig(filename="C:\\Users\\chvatdi00\\Desktop\\wrapper.log", level=logging.DEBUG)

try:
    logging.info(str('initialdir: ') + str(config.initialdir))
    logging.info(str('wrapper_logfile: ') + str(config.wrapper_logfile))
    logging.info(str('oildata_logfile:') + str(config.oildata_logfile))
    logging.info(str('templates:') + str(config.xltemplates))
    logging.info(str('tempfile:') + str(config.tempfile))
    logging.info(str('outfile:') + str(config.outfile))

    # this script is used on windows to wrap shortcuts so that they are executed within an environment
    #   It only sets the appropriate prefix PATH entries - it does not actually activate environments


    #from menuinst.knownfolders import FOLDERID, get_folder_path, PathNotFoundException

    # call as: python cwp.py PREFIX ARGs...

    prefix = sys.argv[1]
    args = sys.argv[2:]

    new_paths = pathsep.join([prefix,
                            join(prefix, "Library", "mingw-w64", "bin"),
                            join(prefix, "Library", "usr", "bin"),
                            join(prefix, "Library", "bin"),
                            join(prefix, "Scripts")])
    env = os.environ.copy()
    env['PATH'] = new_paths + pathsep + env['PATH']
    env['CONDA_PREFIX'] = prefix

    # documents_folder, exception = get_folder_path(FOLDERID.Documents)
    # if exception:
    #     documents_folder, exception = get_folder_path(FOLDERID.PublicDocuments)
    # if not exception:
    #     os.chdir(documents_folder)

    logging.info(f"Wrapper.py called.")
    logging.debug(f"Current Dir: {os.getcwd()}")
    logging.debug(f"CONDA_PREFIX: {env['CONDA_PREFIX']}")
    logging.debug(f"PATH: {env['PATH']}")

    logging.info(sys.argv)

    #workdir = "\\".join(sys.argv[-1].split('\\')[:-1])
    workdir = os.path.dirname(sys.argv[-1]).replace('\\','/')
    logging.debug(f"Work Dir: {workdir}")
    os.chdir(workdir)

    logging.debug(f"subprocess.call({args}, env=env)")
    sys.exit(subprocess.call(args, env=env))
except SystemExit as e:
    logging.info('System exited as intended.')
except Exception as e:
    logging.exception(str(e))
