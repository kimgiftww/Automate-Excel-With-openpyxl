keep_going = True
while keep_going:
    try:
        pyautogui.moveTo(1700, 1000)
        pyautogui.moveTo(800, 1000)
        pyautogui.click()
        pyautogui.click()
        pyautogui.write(text_block, interval=0)
        old_text_block = text_block
        search_string = old_text_block[-10:]

        element_words = browser.find_element_by_id('words')
        soup = BeautifulSoup(element_words.get_attribute('innerHTML'), 'html.parser')
        words = soup.find_all('div', 'word')
        text_block = [word.text for word in words]
        text_block = ' '.join(text_block)
        text_blcok = text_block[text_block.index(search_string):].replace(search_string, '')
    except:
        keep_going = False
