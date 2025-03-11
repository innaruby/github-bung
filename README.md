def locate_image_opencv_multiscale(image_paths, threshold=0.8, scales=np.linspace(0.5, 2.0, 30)):
            try:
                # Ensure image_paths is a list
                if isinstance(image_paths, str):
                    image_paths = [image_paths]

                # Take a screenshot
                screenshot = pyautogui.screenshot()
                screenshot_rgb = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                screenshot_gray = cv2.cvtColor(screenshot_rgb, cv2.COLOR_BGR2GRAY)

                for image_path in image_paths:
                    print(f"Trying to read image: {image_path}")
                    # Load the template image
                    template = cv2.imread(image_path, cv2.IMREAD_COLOR)
                    if template is None:
                        print(f"Failed to read image: {image_path}")
                        continue
                    template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

                    best_match = None
                    best_score = threshold

                    for scale in scales:
                        # Resize the template image
                        resized_template = cv2.resize(template_gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
                        tH, tW = resized_template.shape[:2]

                        # Skip if the template size is larger than the screenshot
                        if tH > screenshot_gray.shape[0] or tW > screenshot_gray.shape[1]:
                            continue

                        # Perform template matching
                        result = cv2.matchTemplate(screenshot_gray, resized_template, cv2.TM_CCOEFF_NORMED)
                        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

                        # Check if we found a better match
                        if max_val > best_score:
                            best_score = max_val
                            best_match = (max_loc, tW, tH)

                    if best_match:
                        max_loc, tW, tH = best_match
                        center_x = int(max_loc[0] + tW / 2)
                        center_y = int(max_loc[1] + tH / 2)
                        return center_x, center_y

                print(f"No match found for any of the images: {image_paths}")
                return None

            except Exception as e:
                print(f"Failed to locate image using OpenCV: {e}")
                return None
def paste_data_to_excel(button_paths):
            try:
                time.sleep(15)
                button_location = locate_image_opencv_multiscale(button_paths)
                if button_location:
                    pyautogui.click(button_location)
                pyautogui.hotkey('ctrl', 'home')
                time.sleep(2)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(4)
                pyautogui.hotkey('ctrl', 's')
                time.sleep(8)
                pyautogui.hotkey('alt', 'f4')
            except Exception as e:
                print(f"An error occurred during paste operation: {e}")

        def click_below_image(image_paths, offset_y=30):
            try:
                image_location = locate_image_opencv_multiscale(image_paths)
                if image_location:
                    x, y = image_location
                    click_position = (x, y + offset_y)
                    pyautogui.moveTo(click_position[0], click_position[1], duration=1)
                    pyautogui.click()
                    print(f"Clicked at position: {click_position}")
                else:
                    print(f"None of the images were found on the screen: {image_paths}")
            except Exception as e:
                print(f"An error occurred: {e}")

     all these functions  must be capable of cases handling where sometimes a single image may be passed on to these functions and othertimes a list of images for example 


 datenart_button_path =   r'U:\datenart_button.png'
        datenart_button_path1 =   r'U:\datenart_button_1.png'
        datenart_button_path2 =   r'U:\datenart_button_2.png'
        datenart_button_path3 =   r'U:\datenart_button_3.png'
        









image_pathsz = [datenart_button_path,  datenart_button_path1 , datenart_button_path2,datenart_button_path3]
                click_below_image(image_pathsz)
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.press('backspace')
                pyautogui.write('I8', interval=0.1)

                


                
