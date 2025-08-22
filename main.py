"""
Universal Excel File Downloader
Downloads Excel files from any website automatically - NO HARDCODED LINKS
"""

import time
import os
from pathlib import Path
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def download_excel_from_website(url, download_folder="./downloads", search_keywords=None):
    """
    Universal Excel file downloader - works with any website
    
    Args:
        url: The website URL to visit
        download_folder: Where to save downloaded files
        search_keywords: Optional list of keywords to help identify the right download
                        (e.g., ['debt', 'statistics'] for GFB site)
    
    Returns:
        Path to downloaded file or None if failed
    """
    
    # Setup
    Path(download_folder).mkdir(exist_ok=True)
    
    # Default search strategies for finding Excel download buttons
    if search_keywords is None:
        search_keywords = []
    
    # Chrome setup
    options = uc.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--headless")
    
    prefs = {
        "download.default_directory": str(Path(download_folder).absolute()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    
    try:
        driver = uc.Chrome(options=options)
    except Exception as e:
        print(f"Failed with options, trying basic setup: {e}")
        options = uc.ChromeOptions()
        options.add_argument("--headless")
        driver = uc.Chrome(options=options)
    
    try:
        print(f"Going to website: {url}")
        driver.get(url)
        
        # Wait for page to fully load
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
        except:
            pass
        
        time.sleep(5)
        
        print("Looking for Excel download buttons...")
        
        # Universal download button detection strategies
        download_buttons = find_excel_download_buttons(driver, search_keywords)
        
        if not download_buttons:
            print("❌ No Excel download buttons found")
            return None
        
        # Try each found button until one works
        for i, button_info in enumerate(download_buttons):
            print(f"Trying download option {i+1}: {button_info['description']}")
            
            result = attempt_download(driver, button_info, download_folder, url)
            if result:
                return result
        
        print("❌ All download attempts failed")
        return None
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return None
        
    finally:
        try:
            driver.quit()
        except:
            pass

def find_excel_download_buttons(driver, search_keywords):
    """
    Find all possible Excel download buttons on the page
    
    Returns:
        List of button information dictionaries
    """
    buttons = []
    
    # Strategy 1: Direct Excel file links
    try:
        elements = driver.find_elements(By.XPATH, "//a[contains(@href, '.xlsx') or contains(@href, '.xls')]")
        for element in elements:
            href = element.get_attribute('href')
            text = element.text.strip()
            title = element.get_attribute('title') or ""
            
            buttons.append({
                'element': element,
                'url': href,
                'description': f"Direct Excel link: {text or title or 'No text'}",
                'method': 'direct_link',
                'relevance_score': calculate_relevance(text + " " + title + " " + href, search_keywords)
            })
    except Exception as e:
        print(f"Strategy 1 failed: {e}")
    
    # Strategy 2: Buttons/links containing download-related text
    download_texts = ['download', 'Download', 'DOWNLOAD', 'excel', 'Excel', 'EXCEL', 'xlsx', 'XLSX']
    for text_pattern in download_texts:
        try:
            elements = driver.find_elements(By.XPATH, f"//a[contains(text(), '{text_pattern}')]")
            for element in elements:
                href = element.get_attribute('href') or ""
                if '.xlsx' in href or '.xls' in href or not href:  # Include buttons without href (might be JS)
                    text = element.text.strip()
                    title = element.get_attribute('title') or ""
                    
                    buttons.append({
                        'element': element,
                        'url': href,
                        'description': f"Download text button: {text}",
                        'method': 'text_button',
                        'relevance_score': calculate_relevance(text + " " + title + " " + href, search_keywords)
                    })
        except Exception as e:
            continue
    
    # Strategy 3: Buttons with download-related attributes
    try:
        elements = driver.find_elements(By.XPATH, "//a[contains(@title, 'download') or contains(@title, 'Download') or contains(@title, 'excel') or contains(@title, 'Excel')]")
        for element in elements:
            href = element.get_attribute('href') or ""
            text = element.text.strip()
            title = element.get_attribute('title') or ""
            
            buttons.append({
                'element': element,
                'url': href,
                'description': f"Download attribute button: {title}",
                'method': 'attribute_button',
                'relevance_score': calculate_relevance(text + " " + title + " " + href, search_keywords)
            })
    except Exception as e:
        print(f"Strategy 3 failed: {e}")
    
    # Strategy 4: Look in common download sections
    section_selectors = [
        "//div[contains(@class, 'download')]//a",
        "//section[contains(@class, 'download')]//a", 
        "//div[contains(@id, 'download')]//a",
        "//div[contains(text(), 'Download')]/following-sibling::*//a",
        "//h2[contains(text(), 'Download')]/following-sibling::*//a",
        "//h3[contains(text(), 'Download')]/following-sibling::*//a"
    ]
    
    for selector in section_selectors:
        try:
            elements = driver.find_elements(By.XPATH, selector)
            for element in elements:
                href = element.get_attribute('href') or ""
                text = element.text.strip()
                title = element.get_attribute('title') or ""
                
                if '.xlsx' in href or '.xls' in href or any(word in (text + title).lower() for word in ['excel', 'xlsx', 'download']):
                    buttons.append({
                        'element': element,
                        'url': href,
                        'description': f"Section download: {text or title or 'No text'}",
                        'method': 'section_search',
                        'relevance_score': calculate_relevance(text + " " + title + " " + href, search_keywords)
                    })
        except Exception as e:
            continue
    
    # Remove duplicates and sort by relevance
    unique_buttons = []
    seen_urls = set()
    
    for button in buttons:
        url = button['url'] or str(button['element'])
        if url not in seen_urls:
            seen_urls.add(url)
            unique_buttons.append(button)
    
    # Sort by relevance score (higher is better)
    unique_buttons.sort(key=lambda x: x['relevance_score'], reverse=True)
    
    print(f"Found {len(unique_buttons)} potential download options")
    for i, button in enumerate(unique_buttons[:5]):  # Show top 5
        print(f"  {i+1}. {button['description']} (score: {button['relevance_score']})")
    
    return unique_buttons

def calculate_relevance(text, search_keywords):
    """
    Calculate relevance score based on search keywords
    
    Args:
        text: Text to analyze
        search_keywords: List of keywords to look for
    
    Returns:
        Relevance score (higher is better)
    """
    if not search_keywords:
        return 1  # Default score when no keywords provided
    
    text_lower = text.lower()
    score = 0
    
    for keyword in search_keywords:
        if keyword.lower() in text_lower:
            score += 2  # Keyword match
    
    # Bonus points for common download indicators
    if any(word in text_lower for word in ['download', 'excel', 'xlsx']):
        score += 1
    
    return score

def attempt_download(driver, button_info, download_folder, base_url):
    """
    Attempt to download using the given button information
    
    Returns:
        Path to downloaded file or None if failed
    """
    try:
        element = button_info['element']
        url = button_info['url']
        
        print(f"Attempting: {button_info['description']}")
        
        # Handle potential overlays/cookie banners first
        handle_overlays(driver)
        
        # Method 1: Direct download if we have a URL
        if url and ('.xlsx' in url or '.xls' in url):
            try:
                result = direct_download(driver, url, download_folder, base_url)
                if result:
                    return result
            except Exception as e:
                print(f"Direct download failed: {e}")
        
        # Method 2: Try clicking the element
        try:
            result = click_download(driver, element, download_folder)
            if result:
                return result
        except Exception as e:
            print(f"Click download failed: {e}")
        
        return None
        
    except Exception as e:
        print(f"Download attempt failed: {e}")
        return None

def handle_overlays(driver):
    """Handle cookie banners and other overlays"""
    try:
        overlay_selectors = [
            "//button[contains(text(), 'Accept')]",
            "//button[contains(text(), 'OK')]", 
            "//button[contains(text(), 'Akzeptieren')]",
            "//button[contains(text(), 'Agree')]",
            "//a[contains(text(), 'Accept')]",
            "//div[contains(@class, 'cookie')]//button",
            "//div[contains(@id, 'cookie')]//button"
        ]
        
        for selector in overlay_selectors:
            try:
                buttons = driver.find_elements(By.XPATH, selector)
                for button in buttons:
                    if button.is_displayed():
                        button.click()
                        time.sleep(1)
                        print("Dismissed overlay")
                        return
            except:
                continue
    except:
        pass

def direct_download(driver, download_url, download_folder, base_url):
    """Download file directly using requests"""
    try:
        # Handle relative URLs
        if download_url.startswith('/'):
            from urllib.parse import urljoin
            download_url = urljoin(base_url, download_url)
        
        print(f"Direct download from: {download_url}")
        
        import requests
        
        # Get browser context
        cookies = driver.get_cookies()
        session_cookies = {cookie['name']: cookie['value'] for cookie in cookies}
        user_agent = driver.execute_script("return navigator.userAgent;")
        
        headers = {
            'User-Agent': user_agent,
            'Referer': base_url,
            'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,*/*'
        }
        
        response = requests.get(download_url, headers=headers, cookies=session_cookies, stream=True)
        response.raise_for_status()
        
        # Get filename dynamically
        filename = get_filename_from_response(response, download_url)
        filepath = os.path.join(download_folder, filename)
        
        # Save file
        with open(filepath, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        
        print(f"✅ SUCCESS! Downloaded: {filename}")
        print(f"📁 Location: {filepath}")
        print(f"📏 Size: {os.path.getsize(filepath)} bytes")
        return filepath
        
    except Exception as e:
        print(f"Direct download error: {e}")
        return None

def click_download(driver, element, download_folder):
    """Try clicking the download element"""
    try:
        # Scroll to element
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
        time.sleep(2)
        
        # Try multiple click methods
        click_methods = [
            lambda: element.click(),
            lambda: driver.execute_script("arguments[0].click();", element),
            lambda: driver.execute_script("arguments[0].dispatchEvent(new MouseEvent('click', {bubbles: true}));", element)
        ]
        
        for i, click_method in enumerate(click_methods):
            try:
                click_method()
                print(f"Clicked using method {i+1}")
                
                # Wait and check for download
                result = wait_for_download(download_folder)
                if result:
                    return result
                    
            except Exception as e:
                print(f"Click method {i+1} failed: {e}")
                continue
        
        return None
        
    except Exception as e:
        print(f"Click download error: {e}")
        return None

def get_filename_from_response(response, url):
    """Extract filename from response or URL"""
    try:
        # Try Content-Disposition header
        if 'content-disposition' in response.headers:
            cd = response.headers['content-disposition']
            if 'filename=' in cd:
                filename = cd.split('filename=')[1].strip('"\'')
                if filename.endswith(('.xlsx', '.xls')):
                    return filename
        
        # Extract from URL
        from urllib.parse import urlparse
        parsed_url = urlparse(url)
        filename = os.path.basename(parsed_url.path)
        
        if filename and filename.endswith(('.xlsx', '.xls')):
            return filename
        
        # Generate timestamped filename
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"excel_download_{timestamp}.xlsx"
        
    except:
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"excel_download_{timestamp}.xlsx"

def wait_for_download(download_folder, timeout=60):
    """Wait for file to be downloaded"""
    try:
        initial_files = set(os.listdir(download_folder)) if os.path.exists(download_folder) else set()
        
        for attempt in range(timeout // 5):
            time.sleep(5)
            
            if os.path.exists(download_folder):
                current_files = set(os.listdir(download_folder))
                new_files = current_files - initial_files
                
                excel_files = [f for f in new_files if f.endswith(('.xlsx', '.xls')) and not f.endswith('.crdownload')]
                
                if excel_files:
                    filename = excel_files[0]  # Get first new Excel file
                    filepath = os.path.join(download_folder, filename)
                    print(f"✅ SUCCESS! Downloaded: {filename}")
                    return filepath
            
            print(f"Waiting for download... {attempt + 1}/{timeout // 5}")
        
        return None
        
    except Exception as e:
        print(f"Wait for download error: {e}")
        return None

# Example usage functions for different websites
def download_gfb_file():
    """Download German Federal Borrowing file"""
    return download_excel_from_website(
        url="https://www.deutsche-finanzagentur.de/en/federal-funding/debt-statistics/gross-borrowing",
        download_folder="./gfb_downloads",
        search_keywords=['debt', 'statistics', 'borrowing', 'schuldenbericht']
    )

def download_from_custom_site(url, keywords=None):
    """Download from any custom website"""
    return download_excel_from_website(
        url=url,
        download_folder="./downloads", 
        search_keywords=keywords
    )

if __name__ == "__main__":
    # Example: Download GFB file
    print("=== German Federal Borrowing Download ===")
    result = download_gfb_file()
    
    if result:
        print(f"File saved to: {result}")
    else:
        print("Download failed")
    
    # Example: Download from any other website
    # result = download_from_custom_site("https://example.com", ['data', 'report'])