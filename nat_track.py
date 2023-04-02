from pathlib import Path
import os
import winreg
import requests
from lxml import etree
from colorama import Fore, Back, Style
from decimal import Decimal
import traceback
import subprocess

VATSYS_MAPS_PATH_RELATIVE = r'vatSys Files\Profiles\gaats-gander-shanwick-dataset\Maps'
VATSYS_PROFILE_PATH_RELATIVE = r'vatSys Files\Profiles\gaats-gander-shanwick-dataset'

NATS_API_URL = 'https://tracks.ganderoceanic.ca/data'

DEFAULT_FILENAME = 'NAT_TRACK.XML'


DEFAULT_MAP_ATTRIBUTES = {
    'Type'             : 'System',
    'Name'             : 'Tracks',
    'Priority'         : '1',
    'CustomColourName' : 'OffWhite'
}

DEFAULT_POLY_ATTRIBUTES = {
    'Type' : 'Line'
}

def error(error_message: str):
    print(Fore.WHITE + Back.RED + 'ERROR:' + Style.RESET_ALL + ' ' + error_message)

def log(log_message: str):
    print(Fore.WHITE + Back.GREEN + 'LOG:' + Style.RESET_ALL + ' ' + log_message)

def clean(coord):
    if len(coord.split('.')[0]) > 3:
        new_coord = coord[0:2]
        return float(f"{new_coord}.5")
    else:
        return coord

def find_vatsys_maps_dir():
    try:
        home_path = r'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, home_path, access=winreg.KEY_READ)
        key_val, _ = winreg.QueryValueEx(key, 'Personal')
        winreg.CloseKey(key)
        full_dir = Path(key_val, VATSYS_MAPS_PATH_RELATIVE)
        if os.path.exists(full_dir):
            return full_dir
    except:
        pass

    try:
        full_dir = Path(str(Path.home()), 'Documents', VATSYS_MAPS_PATH_RELATIVE)
        if os.path.exists(full_dir):
            return full_dir
    except:
        pass

    return None

def find_vatsys_exec():

    full_path = Path(os.environ['ProgramFiles(x86)'], 'vatSys', 'bin', 'vatSys.exe')
    if os.path.exists(full_path):
        return full_path

    full_path = Path(os.environ['ProgramW6432'], 'vatSys', 'bin', 'vatSys.exe')
    if os.path.exists(full_path):
        return full_path
    

    drives = [ chr(x) + ":" for x in range(65,91) if os.path.exists(chr(x) + ":") ]
    for letter in drives:
        for directory in os.listdir(f'{letter}/'):
            try:
                for folders in os.listdir(os.path.join(f'{letter}/', directory)):
                    if folders == 'bin':
                        if 'vatSys.exe' in (os.listdir(os.path.join(f'{letter}/', directory,folders))):
                            return (os.path.join(f'{letter}/', directory, folders, 'vatSys.exe'))
            except:
                pass

    return None


def make_base_map_xml(map_attributes: dict[str, str] = DEFAULT_MAP_ATTRIBUTES) -> tuple[etree.Element, etree.Element]:
    maps_root = etree.Element('Maps')
    map = etree.SubElement(maps_root, 'Map')

    for attribute, val in map_attributes.items():
        map.set(attribute, val)
    return (maps_root, map)

def coord_to_str(coord):
    og_coord = coord
    coord = og_coord.split('|')[0].split('/')

    coord_list = []
    for x in coord:
        
        x = clean(x)

        d = Decimal(str(x))
        integer_part = int(d) - 360 if int(d) > 180 else int(d)
        integer_part_str = str(integer_part)
        fractional_part = d % 1

        if integer_part < 0:
            integer_part_str = f"{integer_part : 04d}"

        fractional_part_str = f'{fractional_part:.3f}'.lstrip('0').lstrip('-0')
        leader = '+' if integer_part > 0 else ''
        coord_list.append(leader + integer_part_str + fractional_part_str)

    return f"{''.join(coord_list)}|{og_coord.split('|')[1]}"

def conversion_func(coord) -> str:
    if coord.find('|') != -1 and coord.split('|')[0].find('/') != -1:
        return coord_to_str(coord)

    return coord
    


def make_poly_xml(track_line: list[list[float]]) -> etree.Element:
    poly_element = etree.Element(DEFAULT_POLY_ATTRIBUTES['Type'])
    point_element = etree.SubElement(poly_element, 'Point')

    point_strings = [conversion_func(coord) for coord in track_line]
    clean_point_strings = []
    for ps in point_strings:
        if ps.find('|') != -1:
            clean_point_strings.append(ps.split('|')[0])
        else:
            clean_point_strings.append(ps)
    point_element.text = '/'.join(clean_point_strings)

    return poly_element

def make_label_xml(series_id_str, first_fix) -> etree.Element:
    label_element = etree.Element('Label')
    
    point_element = etree.SubElement(label_element, 'Point')
    point_element.set('Name', series_id_str)
    
    if first_fix.find('|') != -1:
        first_fix = first_fix.split('|')[1]
    point_element.text = first_fix

    return label_element

def run(vatsys_maps_dir: str, output_filename: str):
    log('Starting')
    
    try:
        r = requests.get(NATS_API_URL)
        tracks_json = r.json()
        log('Nat tracks fetched.')
    except Exception as e:
        error('could not fetch tracks')
        traceback.print_exc()
        exit_error(find_vatsys_exec())

    try:
        maps_root, map_element = make_base_map_xml()

        muff_poly_coords = []
        airspace = etree.Element('Airspace')
        intersections = etree.SubElement(airspace, 'Intersections')
        airways = etree.SubElement(airspace, 'Airways')
        for tracks in tracks_json:
            airway = etree.SubElement(airways, 'Airway',Name=f"NAT{tracks['id']}")

            for track in tracks['route']:
                if track['name'].find('/') != -1:
                    muff_poly_coords.append(f"{track['latitude']}/{track['longitude']}|N{track['name'].replace('/','W')}")
                else:
                    muff_poly_coords.append(track['name'])

            poly_xml = make_poly_xml(muff_poly_coords)
            label_xml = make_label_xml(tracks['id'], muff_poly_coords[0])
            map_element.append(poly_xml)
            map_element.append(label_xml)

            point_strings = [conversion_func(coord) for coord in muff_poly_coords]
            clean_point_strings = []
            for ps in point_strings:
                if ps.find('|') != -1:
                    clean_point_strings.append(ps.split('|')[1])
                    point = etree.SubElement(intersections, 'Point',Name=ps.split('|')[1], Type='Fix')
                    point.text = ps.split('|')[0]
                else:
                    clean_point_strings.append(ps)

            airway.text = '/'.join(clean_point_strings)




            muff_poly_coords.clear()

        path = Path(Path(str(Path.home()), 'Documents', VATSYS_PROFILE_PATH_RELATIVE), 'Airspace.xml')
        etree.ElementTree(airspace).write(path, encoding='UTF-8', xml_declaration=True, pretty_print=True)
        
    except Exception:
        error('could not form XML')
        traceback.print_exc()
        exit_error(find_vatsys_exec())
    
    # Write output XML
    try:
        path = Path(vatsys_maps_dir, output_filename)
        etree.ElementTree(maps_root).write(path, pretty_print=True)
        log('Saved XML file')
    except:
        error('Could not save XML!')
        traceback.print_exc()
        exit()

def exit_error(exec_path):
    if exec_path == False:
        input('Error could not find vatsys. Open manually instead.')
        exit()
    else:
        input('Error. Press ENTER to open vatSys anyway.')
        subprocess.Popen([exec_path])
        exit()


if __name__ == '__main__':

    maps_dir = find_vatsys_maps_dir()
    if maps_dir is None:
        error('could not find suitable vatSys Maps folder for Gander/Shanwick profile')
        exit_error(find_vatsys_exec())
    
    run(maps_dir, DEFAULT_FILENAME)
    
    exec_path = find_vatsys_exec()
    if exec_path is None:
        error('could not find suitable vatSys executable')
        exit_error(False)
    else:
        log(f'opening vatSys executable at {exec_path}')
        subprocess.Popen([exec_path])
        exit()