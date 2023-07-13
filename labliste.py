#!/usr/bin/python3

import sys
import re
import os
import csv
import shutil
import time

str_anrede = 'Anrede'
str_titel = 'Titel'
str_vorname = 'Vorname'
str_nachname = 'Nachname'
str_mitglieds_nr = 'Mitglieds-Nr'
str_zusatzadresse = 'Zusatzadresse'
str_anz_labyrinth = 'AnzLabyrinth'
str_land = 'Land'
required_in=[str_mitglieds_nr, str_anrede, str_vorname, str_nachname, 'Straße', 'PLZ', 'Ort', str_land]
outheader=[str_mitglieds_nr, 'ADR_Z1', 'ADR_Z2', 'ADR_Z3', 'PLZ', 'Ort', str_land, 'Straße', str_anz_labyrinth]

n_error = 0
logfile = None
loglines = []

def err_exit(msg):
    print( msg, file=sys.stderr )
    sys.exit( 1 )

def warn(msg):
    print( msg, file=sys.stderr )

def logerror(msg):
    global n_error
    n_error += 1
    log( 'FEHLER: '+msg )

def printerr(msg):
    print( msg, file=sys.stderr )

def loginfo(msg):
    log( 'INFO: '+msg )

def log(msg):
    printerr( msg )
    loglines.append( msg )

def is_20jj_n(s):
    return re.match( '^20[0-9]{2}-[1-9]$', s )

def read_config(filename):
    file = open( filename, encoding='utf-8-sig' )
    reader = csv.DictReader( file, delimiter=';' )
    config = []
    for line in reader:
        config.append( line )
    return config


def labliste(dir):
    if not os.path.exists( dir ):
        err_exit( dir+" existiert nicht." )
    config = read_config( os.path.join( dir, "konfiguration.csv" ) )
    warning = False
    for cfg in config:
        rv = cfg["RV-KUERZEL"]
        subdir = os.path.join( dir, rv )
        entries = os.listdir( subdir )
        if len( entries ) == 0:
            warn( subdir+" ist leer" )
            warning = True
        elif len( entries ) > 1:
            warn( subdir+" enthält mehr als eine Datei." )
            warning = True
    if warning:
        sys.exit( 1 )
        
    csvrows = []
    addr_total = 0
    for cfg in config:
        rv = cfg["RV-KUERZEL"]
        subdir = os.path.join( dir, rv )
        entries = os.listdir( subdir )
        filename = os.path.join( subdir, entries[0] )
        encoding = 'utf-8-sig'
#        if line == 'HE':
#            encoding = 'iso-8859-1'
        csv_file = open( filename, encoding=encoding )
        reader = csv.DictReader( csv_file, delimiter=';' )
        missing_required = []
        for field in required_in:
            if not field in reader.fieldnames:
                missing_required.append(field)
            if missing_required:
                err_exit( 'Folgende Felder fehlen in '+filename+': '+str( missing_required ) )
        
        addr_no = 0
        mgl_generated = 0
        found_pruef_mgl = False
        for inrow in reader:
            addr_no += 1
            mitglieds_nr = inrow[str_mitglieds_nr]
            if len( mitglieds_nr ) == 0:
                mgl_generated += 1
                mitglieds_nr = "{:06}".format( mgl_generated )
            elif len( mitglieds_nr) != 6:
                logerror( 'Fehlerhafte Mitglieds-Nr in '+filename+'/Adressnummer '+str( addr_no ) )
            if cfg["PRUEF_MGLNR"] != '' and cfg["PRUEF_MGLNR"] == mitglieds_nr:
                found_pruef_mgl = True
            mitglieds_nr = mitglieds_nr # TODO: line+mitglieds_nr
            anrede = inrow[str_anrede]
            if anrede == '':
                adr_z1 = inrow[str_nachname]
                adr_z2 = inrow[str_vorname]
            else:
                adr_z1 = anrede if anrede != 'Herr' else 'Herrn' # TODO: Das ist veraltet
                adr_z2 = inrow[str_vorname]+' '+inrow[str_nachname]
                if str_titel in reader.fieldnames and inrow[str_titel] != '': adr_z2 = inrow[str_titel]+' '+adr_z2
            adr_z3 = inrow[str_zusatzadresse] if str_zusatzadresse in reader.fieldnames else ''
            plz = inrow['PLZ']
            land = inrow[str_land]
            if land.upper() == 'DEUTSCHLAND':
                land = ''
            if land == '' and not re.match( '^[0-9]{5}$', plz ):
                logerror( 'Fehlerhafte PLZ in '+filename+'/Adressnummer '+str( addr_no ) )
            if str_anz_labyrinth in reader.fieldnames:
                anz_labyrinth = inrow[str_anz_labyrinth]
                if anz_labyrinth == '':
                    anz_labyrinth = '1'
                elif not re.match( '^[0-9]{1,3}$', anz_labyrinth ):
                    logerror( 'Fehlerhafter Wert in Spalte '+str_anz_labyrinth+' in '+infn+'/Adressnummer '+str( addr_no ) )
            else:
                anz_labyrinth = '1'
            outrow = [mitglieds_nr, adr_z1, adr_z2, adr_z3, plz, inrow['Ort'], land, inrow['Straße'], anz_labyrinth]
            csvrows.append( outrow )
        loginfo( str( addr_no ).rjust(4)+' Adresse(n) aus '+filename+' gelesen' )
        addr_total += addr_no
        if addr_no < int( cfg["ADDR_MIN"] ) or addr_no > int( cfg["ADDR_MAX"] ):
            logerror( 'Anzahl der Adressen nicht im Bereich '+cfg["ADDR_MIN"]+"-"+cfg["ADDR_MAX"] )
        if cfg["PRUEF_MGLNR"] != '' and not found_pruef_mgl:
            logerror( 'Mitgliedsnummer '+cfg["PRUEF_MGLNR"]+' nicht gefunden' )
        if mgl_generated > 0:
            loginfo( '=> '+str( mgl_generated )+' Mitglieds-Nr erzeugt' )
        csv_file.close()
    if n_error == 0:
        loginfo( '====' )
        loginfo( str( addr_total ).rjust(4)+' Adresse(n) insgesamt' )
        tstamp = time.strftime( '%Y-%m-%d_%H-%M-%S' )

        logfn = os.path.join( dir, "ZIEL", "printec_"+tstamp+".log" )
        csvfn = os.path.join( dir, "ZIEL", "printec_"+tstamp+".csv" )

        logfile = open( logfn, 'w', encoding='UTF-8' )
        csvfile = open( csvfn, 'w', encoding='UTF-8' )

        csvwriter = csv.writer( csvfile, delimiter=';' )
        csvwriter.writerow( outheader )
        for csvrow in csvrows:
            csvwriter.writerow( csvrow )

        for logline in loglines:
            logfile.write( logline+'\n' )
        csvfile.close()
        logfile.close()
        printerr( '' )
        printerr( 'Das Verzeichnis '+dir+' konnte verarbeitet werden, es wurden folgende Ausgabedateien erzeugt:' )
        printerr( '  '+csvfn )
        printerr( '  '+logfn )
    else:
        printerr('Es wurden keine Ausgabedateien erzeugt, da mindestens ein Fehler aufgetreten ist.')
        sys.exit(1)

def verzeichnisse_erstellen(dir):
    if os.path.exists( dir ):
        err_exit( dir+" existiert schon." )
    config = read_config( "konfiguration-neu.csv" )
    os.mkdir( dir )
    for subdir in config:
        os.mkdir( os.path.join( dir, subdir ) )
    os.mkdir( os.path.join( dir, "ZIEL" ) )
    shutil.copyfile( "konfiguration-neu.csv", os.path.join( dir, "konfiguration.csv" ) )

def main():
    if len( sys.argv ) == 2 and is_20jj_n( sys.argv[1] ):
        labliste( sys.argv[1] )
    elif len( sys.argv ) == 3 and sys.argv[1] == 'erstellen' and is_20jj_n( sys.argv[2] ):
        verzeichnisse_erstellen( sys.argv[2] )
    else:
        print( 'Aufruf: '+sys.argv[0]+' [erstellen] 20jj-n', file=sys.stderr )
        sys.exit( 2 )

main()
