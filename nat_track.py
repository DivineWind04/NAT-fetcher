from pathlib import Path
import os
import winreg
import requests
from lxml import etree
import argparse
from colorama import Fore, Back, Style
from decimal import Decimal
import traceback

VATSYS_MAPS_PATH_RELATIVE = r'vatSys Files\Profiles\ATOP Oakland\Maps'
VATSYS_PROFILE_PATH_RELATIVE = r'vatSys Files\Profiles\ATOP Oakland'

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

def exit_with_wait():
    input('Press enter key to exit...')
    exit()

def find_vatsys_maps_dir():
    # First we will try the registry method
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

    # If we failed to find the folder via registry, try with Path method
    try:
        full_dir = Path(str(Path.home()), 'Documents', VATSYS_MAPS_PATH_RELATIVE)
        if os.path.exists(full_dir):
            return full_dir
    except:
        pass
    
    # If we are here, we haven't found anything, so return None
    return None

# def find_vatsys_exec():

#     Returns:
#         A string with the path to the Maps folder if found. None if not found.
#     """
#     # Try the x86 folder first
#     full_path = Path(os.environ['ProgramFiles(x86)'], 'vatSys', 'bin', 'vatSys.exe')
#     if os.path.exists(full_path):
#         return full_path

#     # Next try the regular Program Files folder
#     full_path = Path(os.environ['ProgramW6432'], 'vatSys', 'bin', 'vatSys.exe')
#     if os.path.exists(full_path):
#         return full_path
    
#     # Return none if both fail
#     return None



def make_base_map_xml(map_attributes: dict[str, str] = DEFAULT_MAP_ATTRIBUTES) -> tuple[etree.Element, etree.Element]:
    maps_root = etree.Element('Maps')
    map = etree.SubElement(maps_root, 'Map')

    for attribute, val in map_attributes.items():
        map.set(attribute, val)
    return (maps_root, map)

def coord_to_str(coord):
    coord = coord.split('/')
    coord_list = []
    for x in coord:
        d = Decimal(str(x))
        integer_part = int(d) - 360 if int(d) > 180 else int(d)
        integer_part_str = str(integer_part)
        fractional_part = d % 1

        if integer_part < 0:
            integer_part_str = f"{integer_part : 04d}"

        fractional_part_str = f'{fractional_part:.3f}'.lstrip('0').lstrip('-0')
        leader = '+' if integer_part > 0 else ''
        coord_list.append(leader + integer_part_str + fractional_part_str)
    
    return ''.join(coord_list)

def conversion_func(coord) -> str:
    if coord.find('/') != -1:
        return coord_to_str(coord)
    return coord
    


def make_poly_xml(track_line: list[list[float]]) -> etree.Element:
    # Create the empty elements
    poly_element = etree.Element(DEFAULT_POLY_ATTRIBUTES['Type'])
    point_element = etree.SubElement(poly_element, 'Point')

    # Create all the point strings from the poly coords and add inside <Point> element
    point_strings = [conversion_func(coord) for coord in track_line]
    point_element.text = '/'.join(point_strings)
    log('created polygon with ISO 6709 coordinates %s' % point_strings)

    return poly_element

def make_label_xml(series_id_str, first_fix) -> etree.Element:
    # Create Label element
    label_element = etree.Element('Label')
    
    point_element = etree.SubElement(label_element, 'Point')
    point_element.set('Name', series_id_str)
    point_element.text = first_fix
    log('created label with ISO 6709 coordinates and label %s' % (series_id_str))

    return label_element

def run(vatsys_maps_dir: str, output_filename: str):
    log('running with output location %s' % Path(vatsys_maps_dir, output_filename))
    
    # Fetch Tracks from API
    try:
        r = requests.get(NATS_API_URL)
        tracks_json = r.json()
        log('fetched XML from %s' % NATS_API_URL)
    except Exception as e:
        error('could not fetch Tracks from API')
        traceback.print_exc()
        exit_with_wait()

    # Make the XML
    try:
        # Make the base <Maps> and <Map> element
        maps_root, map_element = make_base_map_xml()
        # Iterate over each NAT XML feature and make the Poly xml (<Infill> or <Line>). Add to <Map>
        muff_poly_coords = []
        airspace = etree.Element('Airspace')
        airways = etree.SubElement(airspace, 'Airways')
        for tracks in tracks_json:
            airway = etree.SubElement(airways, 'Airway',Name=f"NAT{tracks['id']}")

            for track in tracks['route']:
                if track['name'].find('/') != -1:
                    muff_poly_coords.append(f"{track['latitude']}/{track['longitude']}")
                else:
                    muff_poly_coords.append(track['name'])

            poly_xml = make_poly_xml(muff_poly_coords)
            label_xml = make_label_xml(tracks['id'], muff_poly_coords[0])
            map_element.append(poly_xml)
            map_element.append(label_xml)

            point_strings = [conversion_func(coord) for coord in muff_poly_coords]
            airway.text = '/'.join(point_strings)

            muff_poly_coords.clear()

        path = Path(Path(str(Path.home()), 'Documents', VATSYS_PROFILE_PATH_RELATIVE), 'Airspace.xml')
        etree.ElementTree(airspace).write(path, encoding='UTF-8', xml_declaration=True, pretty_print=True)
        
    except Exception:
        error('could not form XML')
        traceback.print_exc()
        exit_with_wait()
    
    # Write output XML
    try:
        path = Path(vatsys_maps_dir, output_filename)
        etree.ElementTree(maps_root).write(path, pretty_print=True)
        log('wrote XML file to %s' % path)
    except:
        error('could not write output file to %s' % path)
        traceback.print_exc()
        exit_with_wait()


if __name__ == '__main__':

    ## Creating the argument parser
    ## TODO: add options for verbosity? or to launch vatSys after? need to implement color option
    parser = argparse.ArgumentParser()
    parser.add_argument('--mapsdir', help="location of vatSys Maps folder for ATOP Oakland profile")
    parser.add_argument('--filename', help="full name of output XML file (including .xml)")
    parser.add_argument('--exec', help="location of vatSys executable")
    parser.add_argument('--color', help="name of vatSys color (from Colours.xml) to use for Tracks")
    args = parser.parse_args()

    # Get profile maps dir from command line first, or do auto. Fail out if we can't find
    maps_dir = args.mapsdir if args.mapsdir is not None else find_vatsys_maps_dir()
    if maps_dir is None:
        error('could not find suitable vatSys Maps folder for ATOP Oakland profile')
        exit_with_wait()
    
    # Get output filename for command line first, or just default
    filename = args.filename if args.filename is not None else DEFAULT_FILENAME

    # We've got the maps_dir and filename now, so we can run
    run(maps_dir, filename)
    
    # Get the vatSys executable to run after
    # exec_path = args.exec if args.exec is not None else find_vatsys_exec()
    # if exec_path is None:
    #     error('could not find suitable vatSys executable')
    #     exit_with_wait()
    # else:
    #     log('opening vatSys executable at %s' % exec_path)
    #     subprocess.Popen([exec_path])
    #     exit_with_wait()
