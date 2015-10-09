#!/usr/bin/env python
from __future__ import print_function, division
from pptx import Presentation
from pptx.util import Inches, Pt
#from pptx.enum.dml import MSO_THEME_COLOR
from pptx.dml.color import RGBColor
from datetime import time as dtime
from datetime import datetime, timedelta
import os
import re
import urllib

# The next three lines automatically
# set the presentation key date to be
# 12Z today
now = datetime.now()
nowdate = now.date()
present_date = datetime.combine(nowdate,dtime(12,0))
# Uncomment the following line to manually specify the
# key date of the presentation
#present_date=datetime(2015,9,10,12)


# Set the directory where the output files will be put
presentation_path = '.'
# Set the path to the model images
model_path='http://www.atmos.washington.edu/~lmadaus/olympex/index.php?init={:%Y%m%d%H}&product={:s}&start={:d}'

# Different default possibilities for the slide layout
layout = {'Title Slide' : 0,
          'Bullet Slide' : 1,
          'Segue' : 2,
          'Side By Side' : 3,
          'Title Alone' : 5,
          'Blank Slide' : 6,
          'Picture with Caption' : 8,
          
         }

# The pattern for the image paths is:
#   'Product name' : (path to file or web address, filename ending )
img_paths = {
             #'IR+500mb' : ('http://www.atmos.washington.edu/images/sat_upr/YYYYMMDDHH00_500mb.gif','500mb.gif',),
             #'Water Vapor' : ('http://www.atmos.washington.edu/images/sat_common/YYYYMMDDHH00_wv.gif','wv.gif'),
             'IR+500mb' : ('http://www.atmos.washington.edu/cgi-bin/latest.cgi?sat_500+-notitle','500mb.gif',),
             'Water Vapor' : ('http://www.atmos.washington.edu/cgi-bin/latest.cgi?wv_common_full+-notitle','wv.gif'),
             'OPC Surface Analys.' : ('http://www.opc.ncep.noaa.gov/P_e_sfc_color.png',None),
             'WRF 500mb Day 0' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxLG_500mb.YYYYMMDDHH.f006.png', 'wrf_500mb_day0.png'),
             'WRF 500mb Day 1' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxLG_500mb.YYYYMMDDHH.f024.png', 'wrf_500mb_day1.png'),
             'WRF 500mb Day 2' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxLG_500mb.YYYYMMDDHH.f048.png', 'wrf_500mb_day2.png'),
             'WRF SLP Day 0' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxLG_surface.YYYYMMDDHH.f006.png', 'wrf_sfc_day0.png'),
             'WRF SLP Day 1' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxLG_surface.YYYYMMDDHH.f024.png', 'wrf_sfc_day1.png'),
             'WRF SLP Day 2' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxLG_surface.YYYYMMDDHH.f048.png', 'wrf_sfc_day2.png'),

             'WRF Melt. Level Day 0' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxSM_melt_level.YYYYMMDDHH.f006.png', 'wrf_melt_level_day0.png'),
             'WRF Melt. Level Day 1' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxSM_melt_level.YYYYMMDDHH.f024.png', 'wrf_melt_level_day1.png'),     
             'WRF Melt. Level Day 2' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxSM_melt_level.YYYYMMDDHH.f048.png', 'wrf_melt_level_day2.png'),     

             'WRF 12hr Prcp Day 0' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxLG_precip12hr.YYYYMMDDHH.f012.png', 'wrf_precip_large_day0.png'),
             'WRF 12hr Prcp (4km) Day 0' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxSM_precip12hr.YYYYMMDDHH.f012.png', 'wrf_precip_small_day0.png'),
             'WRF 12hr Prcp Day 1' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxLG_precip12hr.YYYYMMDDHH.f024.png', 'wrf_precip_large_day1.png'),
             'WRF 12hr Prcp (4km) Day 1' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxSM_precip12hr.YYYYMMDDHH.f024.png', 'wrf_precip_small_day1.png'),
             'WRF 12hr Prcp Day 2' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxLG_precip12hr.YYYYMMDDHH.f048.png', 'wrf_precip_large_day2.png'),
             'WRF 12hr Prcp (4km) Day 2' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxSM_precip12hr.YYYYMMDDHH.f048.png', 'wrf_precip_small_day2.png'),


             'WRF 3hr Prcp Day 0' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxLG_precip3hr.YYYYMMDDHH.f012.png', 'wrf_precip03_large_day0.png'),
             'WRF 3hr Prcp (4km) Day 0' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxSM_precip3hr.YYYYMMDDHH.f012.png', 'wrf_precip03_small_day0.png'),
             'WRF 3hr Prcp Day 1' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxLG_precip3hr.YYYYMMDDHH.f024.png', 'wrf_precip03_large_day1.png'),
             'WRF 3hr Prcp (4km) Day 1' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxSM_precip3hr.YYYYMMDDHH.f024.png', 'wrf_precip03_small_day1.png'),
             'WRF 3hr Prcp Day 2' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxLG_precip3hr.YYYYMMDDHH.f048.png', 'wrf_precip03_large_day2.png'),
             'WRF 3hr Prcp (4km) Day 2' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxSM_precip3hr.YYYYMMDDHH.f048.png', 'wrf_precip03_small_day2.png'),


             'WRF 10m Wind Day 0' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxLG_wssfc.YYYYMMDDHH.f006.png', 'wrf_wssfc_large_day0.png'),
             'WRF 10m Wind (4km) Day 0' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxSM_wssfc.YYYYMMDDHH.f006.png', 'wrf_wssfc_small_day0.png'),
             'WRF 10m Wind Day 1' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxLG_wssfc.YYYYMMDDHH.f024.png', 'wrf_wssfc_large_day1.png'),
             'WRF 10m Wind (4km) Day 1' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxSM_wssfc.YYYYMMDDHH.f024.png', 'wrf_wssfc_small_day1.png'),
             'WRF 10m Wind Day 2' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxLG_wssfc.YYYYMMDDHH.f048.png', 'wrf_wssfc_large_day2.png'),
             'WRF 10m Wind (4km) Day 2' : ('http://www.atmos.washington.edu/~lmadaus/olympex/wrf_plots/YYYYMMDDHH/opxSM_wssfc.YYYYMMDDHH.f048.png', 'wrf_wssfc_small_day2.png'),             
             
             'NWS PacNW Radar' : ('http://www.atmos.washington.edu/~lmadaus/olympex/radar/YYYYMMDDHH00_bref.png','radar.png'),    
             'KUIL Latest Sound.' : ('http://www.atmos.washington.edu/~lmadaus/olympex/soundings/KUIL_YYYYMMDDHH_snd.png','sounding.png'),
             'GFS 500mb Day 2' : ('','gfs_500_day3.gif'),
             'NAEFS SLP and Spread Day 2' : ('http://collaboration.cmc.ec.gc.ca/cmc/ensemble/cartes/data/cartes/PN/CMC_NCEP','naefs_slp_day2.gif'),
             'NAEFS SLP and Spread Day 3' : ('http://collaboration.cmc.ec.gc.ca/cmc/ensemble/cartes/data/cartes/PN/CMC_NCEP','naefs_slp_day3.gif'),
             'NAEFS SLP and Spread Day 4' : ('http://collaboration.cmc.ec.gc.ca/cmc/ensemble/cartes/data/cartes/PN/CMC_NCEP','naefs_slp_day4.gif'),
             'NAEFS SLP and Spread Day 5' : ('http://collaboration.cmc.ec.gc.ca/cmc/ensemble/cartes/data/cartes/PN/CMC_NCEP','naefs_slp_day5.gif'),

             'NAEFS 500mb and Spread Day 2' : ('http://collaboration.cmc.ec.gc.ca/cmc/ensemble/cartes/data/cartes/GZ500/CMC_NCEP','naefs_500mb_day2.gif'),
             'NAEFS 500mb and Spread Day 3' : ('http://collaboration.cmc.ec.gc.ca/cmc/ensemble/cartes/data/cartes/GZ500/CMC_NCEP','naefs_500mb_day3.gif'),
             'NAEFS 500mb and Spread Day 4' : ('http://collaboration.cmc.ec.gc.ca/cmc/ensemble/cartes/data/cartes/GZ500/CMC_NCEP','naefs_500mb_day4.gif'),
             'NAEFS 500mb and Spread Day 5' : ('http://collaboration.cmc.ec.gc.ca/cmc/ensemble/cartes/data/cartes/GZ500/CMC_NCEP','naefs_500mb_day5.gif'),

            }

def build_presentation(present_date):
    """
    Function to build the final presentation
    Inputs:
    present_date --> The key date for the presentation (Datetime object)
         (i.e., the initialization date for the model graphics to grab)
    """
    # Switch to the presentation directory
    basedir = os.getcwd()
    os.chdir(presentation_path)
    
    # Make a new subdirectory for this presentation
    if os.path.exists('./{:%Y%m%d%H}'.format(present_date)):
        os.system('rm ./{:%Y%m%d%H}/*'.format(present_date))
    else:     
        os.mkdir('./{:%Y%m%d%H}'.format(present_date))
    os.chdir('./{:%Y%m%d%H}'.format(present_date))

    # Make a new presentation
    #try:
    #prs = Presentation(os.path.join(basedir,'default.pptx'))
    #except:
    prs = Presentation()
    
    # Make the title slide
    # Choose a title slide layout
    title_slide_layout = prs.slide_layouts[layout['Title Slide']]
    # Add the slide to the presentation
    title_slide = prs.slides.add_slide(title_slide_layout)
    # Change the title
    if present_date.hour < 12:
        title_slide.shapes.title.text = "Evening Weather Update"

    else:  
        title_slide.shapes.title.text = "Morning Weather Briefing"
    # Subtitle is a "placeholder" object
    title_slide.placeholders[1].text = "{:%d %b %Y / %H00Z}\nForecaster Name".format(present_date)

    
    # Make Prev weather bumper
    prs = bumper_slide(prs, 'Past 24 hours', present_date-timedelta(days=1))
    # Blank slide here
    prs = full_summary(prs, 'Summary of Prev. 24 Hours')



    # Make current weather bumper
    prs = bumper_slide(prs, 'Current Weather', present_date)
    # Call the above function to add full-slide images
    # Here for IR+500
    prs = full_slide_image(prs, 'IR+500mb', present_date, width=9, link='http://www.atmos.washington.edu/~ovens/wxloop.cgi?sat_500+/2d/')
    # And for Water Vapor
    prs = full_slide_image(prs, 'Water Vapor', present_date, width=9, link='http://www.atmos.washington.edu/~ovens/wxloop.cgi?wv_common+/48h/')
    # Here for the web-based OPC surface analysis
    prs = full_slide_image(prs, 'OPC Surface Analys.', present_date, link=False)

    # Model verification?    

    
    # Regional Radar
    prs = full_slide_image(prs, 'NWS PacNW Radar', present_date, link='http://www.atmos.washington.edu/~lmadaus/olympex/index.php?product=radar')
    # Current Sounding
    prs = full_slide_image(prs, 'KUIL Latest Sound.', present_date)
    
    # Current airport conditions--> vis, wind dir, wind speed (and at NPOL)
    #prs = full_summary(prs, 'Current Airport Conditions')
    prs = airport_slide(prs, 'Current Airport Conditions')
    
    # Bumper into next 24 hour forecast
    prs = bumper_slide(prs, 'Forecast: Day 0', present_date)
    day0_ftime = present_date + timedelta(hours=6)
    #day1_ftime = day1_ftime.replace(hour=0, minute=0, second=0)
    
   
    # WRF Image --> 500mb Vort
    prs = full_slide_image(prs, 'WRF 500mb Day 0', present_date, day0_ftime, link=model_path.format(present_date,'opxLG_500mb',1))
    
    # WRF SLP
    prs = full_slide_image(prs, 'WRF SLP Day 0', present_date, day0_ftime, link=model_path.format(present_date,'opxLG_surface',1))
    
    # WRF Melting level
    prs = full_slide_image(prs, 'WRF Melt. Level Day 0', present_date, day0_ftime, link=model_path.format(present_date,'opxSM_melt_level',1))
       
    # WRF zoom Precip
    prs = full_slide_image(prs, 'WRF 3hr Prcp (4km) Day 0', present_date, day0_ftime+timedelta(hours=6), link=model_path.format(present_date,'opxSM_precip3hr',1)) 
      
    # WRF zoom 10m Winds
    prs = full_slide_image(prs, 'WRF 10m Wind (4km) Day 0', present_date, day0_ftime, link=model_path.format(present_date,'opxSM_wssfc',1))
 
    # GPM Overpasses
    prs = full_summary(prs, 'GPM Overpasses')

    # Summary
    prs = objectives_slide(prs, 'Day 0 Summary')    
    
    # Bumper into day 1 forecast
    prs = bumper_slide(prs, 'Forecast: Day 1', present_date + timedelta(hours=12))
    day1_ftime = present_date + timedelta(hours=24)
    #day2_ftime = day2_ftime.replace(hour=0, minute=0, second=0)
      
    # WRF Image --> 500mb Vort
    prs = full_slide_image(prs, 'WRF 500mb Day 1', present_date, day1_ftime, link=model_path.format(present_date,'opxLG_500mb',5))
    
    # WRF SLP
    prs = full_slide_image(prs, 'WRF SLP Day 1', present_date, day1_ftime, link=model_path.format(present_date,'opxLG_surface',5))
    
    # WRF Melting level
    prs = full_slide_image(prs, 'WRF Melt. Level Day 1', present_date, day1_ftime, link=model_path.format(present_date,'opxSM_melt_level',13))
    
    # WRF zoom Precip
    prs = full_slide_image(prs, 'WRF 12hr Prcp (4km) Day 1', present_date, day1_ftime, link=model_path.format(present_date,'opxSM_precip12hr',13)) 

  
    # WRF 3hr zoom Precip
    prs = full_slide_image(prs, 'WRF 3hr Prcp (4km) Day 1', present_date, day1_ftime, link=model_path.format(present_date,'opxSM_precip3hr',13)) 
    
    # WRF 10m Winds
    prs = full_slide_image(prs, 'WRF 10m Wind (4km) Day 1', present_date, day1_ftime, link=model_path.format(present_date,'opxSM_wssfc',13))  

    # GPM Overpasses
    prs = full_summary(prs, 'GPM Overpasses')
   
    # Possible objectives
    prs = objectives_slide(prs, 'Day 1 Summary')

    # Bumper into day 2 forecast
    prs = bumper_slide(prs, 'Forecast: Day 2', present_date + timedelta(hours=36))
    day2_ftime = present_date + timedelta(hours=48)

    # WRF Image --> 500mb Vort
    prs = full_slide_image(prs, 'WRF 500mb Day 2', present_date, day2_ftime, link=model_path.format(present_date,'opxLG_500mb',13))
    
    # WRF SLP
    prs = full_slide_image(prs, 'WRF SLP Day 2', present_date, day2_ftime, link=model_path.format(present_date,'opxLG_surface',13))
    
    # WRF Melting level
    prs = full_slide_image(prs, 'WRF Melt. Level Day 2', present_date, day2_ftime, link=model_path.format(present_date,'opxSM_melt_level',37))
    
    # WRF Precip
    prs = full_slide_image(prs, 'WRF 12hr Prcp Day 2', present_date, day2_ftime, link=model_path.format(present_date,'opxLG_precip12hr',13))    

    # WRF zoom Precip
    prs = full_slide_image(prs, 'WRF 12hr Prcp (4km) Day 2', present_date, day2_ftime, link=model_path.format(present_date,'opxSM_precip12hr',37)) 
    
    # WRF 3hr zoom Precip
    prs = full_slide_image(prs, 'WRF 3hr Prcp (4km) Day 2', present_date, day2_ftime, link=model_path.format(present_date,'opxSM_precip3hr',37))     
    
    # WRF 10m Winds
    prs = full_slide_image(prs, 'WRF 10m Wind (4km) Day 2', present_date, day2_ftime, link=model_path.format(present_date,'opxSM_wssfc',37))  

    # GPM Overpasses
    prs = full_summary(prs, 'GPM Overpasses')
   
    # Summary
    prs = objectives_slide(prs, 'Day 2 Summary')





    # Bumper into day 3+ forecast
    prs = bumper_slide(prs, 'Forecast: Day 3+', present_date + timedelta(days=3))
    
    # WRF Image --> 500mb Vort
    #prs = full_slide_image(prs, 'WRF 500mb Day 3', day2_ftime)
    # NAEFS uncertainty   
    prs = full_slide_image(prs, 'NAEFS 500mb and Spread Day 3', present_date, link="https://weather.gc.ca/ensemble/naefs/cartes_e.html")
    prs = full_slide_image(prs, 'NAEFS 500mb and Spread Day 4', present_date, link="https://weather.gc.ca/ensemble/naefs/cartes_e.html")
    prs = full_slide_image(prs, 'NAEFS 500mb and Spread Day 5', present_date, link="https://weather.gc.ca/ensemble/naefs/cartes_e.html")
    # Summary
    prs = objectives_slide(prs, 'Day 3+ Summary')

    
    # Conclusion slide
    prs = full_summary(prs, 'Discussion Summary')
    
    # Save the presentation
    prs.save('wxbriefing_{:%Y%m%d%H}.pptx'.format(present_date))

def get_latest_image(product, present_date, within_hours=12):
    """
    Function to get the latest image from a given product,
    but only if it is within within_hours of present_date
    
    Returns a tuple with:
    
    (string that is the full address of the image,
    date the image is valid (datetime object))
    
    The product name MUST HAVE an entry in the img_paths dictionary,
    otherwise None is returned
    """
    gfs_hours = {3:60,
                 4:84,
                 5:108}
    naefs_hours = {3:72,
                   4:96,
                   5:120}
    
    if product not in img_paths.keys():
        print("Unable to find path to product:", product)
        return None
    # Parse out the info
    path, ext = img_paths[product]
    # If this is a path, replace the starttime with the desired time
    # Given as present date (this MUST be a model start time)
    path = path.replace('YYYYMMDDHH',present_date.strftime('%Y%m%d%H'))    
    
    # This is a web address
    if product in ['GFS 500mb Day 3', 'NAEFS 500mb and Spread Day 3','NAEFS 500mb and Spread Day 4',\
                'NAEFS 500mb and Spread Day 5']:
        recent_file = ext
        if os.path.exists(recent_file):
            os.system('rm -f {:s}'.format(recent_file))
            
        # Replace the DDDD in the path with the current date
        
        if product.startswith('GFS'):
            fhour = gfs_hours[int(product[-1])]
            fdate = present_date.replace(hour=12)
        elif product.startswith('NAEFS'):
            fhour = naefs_hours[int(product[-1])]
            fdate = present_date.replace(hour=0)
            
        path = path + '/{:%Y%m%d%H}_{:03d}.gif'.format(fdate,fhour)
        #print(path)
        try:
            urllib.request.urlretrieve(path, recent_file)
        except AttributeError:
            urllib.urlretrieve(path, recent_file)
            
        fdate = None
    
    elif ext is None:
        # Just download directly from the path
        fdate = None
        recent_file = path.split('/')[-1]
        if os.path.exists(recent_file):
            os.system('rm -f {:s}'.format(recent_file))
            
        try:    
            urllib.request.urlretrieve(path, recent_file)
        except AttributeError:
            urllib.urlretrieve(path, recent_file)
    else:
        if 'WRF' in product:
            fhour = int(path.split('.')[-2][1:])
            fdate = present_date + timedelta(hours=fhour)
            pathsplit = path.split('.')
            pathsplit[-3] = fdate.strftime('%Y%m%d%H')
            path = '.'.join(pathsplit)
        else:
            fdate=None
        #print(path)
        if os.path.exists(ext):
            os.system('rm -f {:s}'.format(ext))
                #print(path)
        try:
            urllib.request.urlretrieve(path, ext)
        except AttributeError:
            urllib.urlretrieve(path, ext)
        #except:
        #   print("FILE NOT FOUND:")
        #   print(path)
        #   exit(1)
            
        recent_file = ext
            
    path = ''


        

    # Now check if the file is close in time to presentation date
    """
    if fdate is not None:
        tdiff = present_date - fdate
        nhours = tdiff.days * 24 + tdiff.seconds/3600.
        if abs(nhours) > within_hours:
            print("{:s} most recent time is {:%H%MZ %d %b %Y}, skipping".format(product, fdate))
            return None
        print("Found {:s} image valid at {:%H%MZ %d %b %Y}".format(product, fdate))
    else:
    """
    print("Found {:s} image".format(product))
    # Return the path
    if path == '':
        return (recent_file, fdate)
    else:
        return ('/'.join((path,recent_file)), fdate)




def full_slide_image(prs,product,present_date, ftime=None, width=None, link=False):
    # Take "product" and make a full-slide image with title out of it
    # Grab the latest image
    results = get_latest_image(product, present_date)
    
 
    #imgpath = product
   
    # Get a blank slide layout and add it to the presentation
    slide_layout = prs.slide_layouts[layout['Title Alone']]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    if width is not None:
        title.top=Inches(7.1)
        title.left = Inches(0)
        title.width=Inches(10)
    else:
        title.top=Inches(3)
        title.left = Inches(7.5)
        title.width=Inches(2.0)
    p = title.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = product
    r.font.size=Pt(40)
    if link:
        hlink = r.hyperlink
        hlink.address = link
    if ftime is not None:
        d = p.add_run()
        d.text = '\n\n' + ftime.strftime('%d %b %HZ')
        d.font.size=Pt(28)

    if results is None:
        # Didn't find the image.  Create slide anyway.
        return prs
    imgpath, imgdate = results   
    # Add the image
    if width is not None:
        left_balanced = (10-width)/2.
        pic = slide.shapes.add_picture(imgpath, left=Inches(left_balanced), top=Inches(0.1), width=Inches(width))
    else:
        pic = slide.shapes.add_picture(imgpath, left=Inches(0), top=Inches(0.0), width=Inches(7))
    return prs



def bumper_slide(prs, title, date):
    # Choose a blank
    slide_layout = prs.slide_layouts[layout['Segue']]
    # Add the slide to the presentation
    slide = prs.slides.add_slide(slide_layout)
    # Change the title
    slide.shapes.title.text = title
    # Subtitle is date
    if 'Current Weather' in title:
        slide.placeholders[1].text = ''
    elif 'Day 0' not in title:
        end = date + timedelta(hours=24)
        slide.placeholders[1].text = "{:%HZ %d %b %Y} through {:%HZ %d %b %Y}".format(date, end)

    else:
        end = date + timedelta(hours=12)
        slide.placeholders[1].text = "Now through {:%HZ %d %b %Y}".format(end)

    return prs



def objectives_slide(prs, title):
    # Choose a blank
    slide_layout = prs.slide_layouts[layout['Side By Side']]
    # Add the slide to the presentation
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title 
    
    return prs

def full_summary(prs, title):
    slide_layout = prs.slide_layouts[layout['Bullet Slide']]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    return prs

def airport_slide(prs, title):
    slide_layout = prs.slide_layouts[layout['Bullet Slide']]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    # Get main text box
    mainbox = slide.placeholders[1]
    tf = mainbox.text_frame
    tf.clear()
    # New paragraph for each station
    stations = ['Paine Field [KPAE]','McChord Field [KTCM]','Hoquiam [KHQM]']
    for s in stations:
        p = tf.add_paragraph()
        run = p.add_run()
        run.text = s + '\n'
        run = p.add_run()
        run.text = '\tWind:\n'
        run = p.add_run()
        run.text = '\tCeiling:\n'
        run = p.add_run()
        run.text = '\tVisibility:'
    tf.margin_bottom=Inches(0.1)
    tf.word_wrap = False

    
    return prs    


if __name__ == '__main__':
    build_presentation(present_date)




