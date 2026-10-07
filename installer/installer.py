# Installe SchoolBot Mail et crée l'icône sur le Bureau.
# Fichier : installer/installer.py  (construit en « Installer SchoolBot Mail.exe » par GitHub)
import ctypes
import os
import shutil
import subprocess
import sys

APP = "SchoolBot Mail"
DEST = r"C:\SchoolBotMail"
DETACHED = 0x00000008 | 0x00000200  # fenêtre indépendante + nouveau groupe de processus
NOWINDOW = 0x08000000


def message(texte, drapeau=0x40):
    """Boîte de dialogue Windows (aucune fenêtre noire, aucun message technique)."""
    ctypes.windll.user32.MessageBoxW(None, texte, APP, drapeau)


def source():
    """Dossier « SchoolBot Mail » situé à côté de cet installateur."""
    base = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
    return os.path.join(base, APP)


def dossier_installation():
    """C:\\SchoolBotMail, ou un dossier personnel si Windows refuse C:\\"""
    try:
        os.makedirs(DEST, exist_ok=True)
        return DEST
    except OSError:
        alternatif = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), APP)
        os.makedirs(alternatif, exist_ok=True)
        return alternatif


def fermer_ancienne_version():
    """Ferme SchoolBot Mail s'il tourne encore, pour pouvoir le remplacer."""
    subprocess.run(["taskkill", "/F", "/IM", APP + ".exe"],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=NOWINDOW)
    import time; time.sleep(1.5)


def copier(src, dest):
    shutil.copytree(src, dest, dirs_exist_ok=True)


def icone_bureau(dest):
    exe = os.path.join(dest, APP + ".exe")
    commande = (
        "$s=(New-Object -ComObject WScript.Shell).CreateShortcut("
        "[Environment]::GetFolderPath('Desktop')+'\\%s.lnk');"
        "$s.TargetPath='%s';$s.WorkingDirectory='%s';$s.Save()" % (APP, exe, dest)
    )
    subprocess.run(["powershell", "-NoProfile", "-Command", commande],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=NOWINDOW)


def lancer(dest):
    subprocess.Popen([os.path.join(dest, APP + ".exe")], cwd=dest, creationflags=DETACHED)


def main():
    src = source()
    if not os.path.isfile(os.path.join(src, APP + ".exe")):
        message(
            "Le programme à installer est introuvable à côté de cet installateur.\n\n"
            "Voici la marche à suivre :\n"
            "   1. Clic droit sur le fichier téléchargé (le .zip)\n"
            "   2. Choisissez « Extraire tout… »\n"
            "   3. Ouvrez le dossier qui vient d'être créé\n"
            "   4. Double-cliquez sur « Installer SchoolBot Mail »\n\n"
            "Ne lancez jamais l'installation directement dans le .zip.",
            0x30,
        )
        return

    dest = dossier_installation()
    if os.path.normcase(src) == os.path.normcase(dest):
        lancer(dest)
        return

    fermer_ancienne_version()
    try:
        copier(src, dest)
    except OSError:
        message(
            "SchoolBot Mail est encore ouvert, impossible de l'installer.\n\n"
            "Fermez la fenêtre SchoolBot Mail (bouton croix, ou Gestionnaire des tâches),\n"
            "puis double-cliquez de nouveau sur « Installer SchoolBot Mail ».",
            0x30,
        )
        return

    icone_bureau(dest)
    message(
        "Installation terminée !\n\n"
        "Une icône « SchoolBot Mail » a été ajoutée sur votre Bureau.\n"
        "Double-cliquez dessus pour démarrer.",
        0x40,
    )
    lancer(dest)


try:
    main()
except Exception as e:
    message("Une erreur est survenue pendant l'installation :\n\n%s\n\n"
            "Écrivez à contact@educlan.org pour être aidé." % e, 0x10)
