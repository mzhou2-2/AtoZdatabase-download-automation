# automate downloading AtoZdatabases data

import pyautogui
import time

# Safety Feature: Move mouse to top-left to abort
pyautogui.FAILSAFE = True

# Short pause after each PyAutoGUI call
pyautogui.PAUSE = 1.0


# Function to move and click
def move_and_click(x, y, duration = 0.35):
    pyautogui.moveTo(x, y, duration=duration)
    pyautogui.click()

# Function to type a number
def type_number(number):
    pyautogui.typewrite(str(number), interval = 0.1)

# Function to wait for user confirmation to start
def wait_for_start():
    input("Press Enter to start the automation...")

# Function to gracefully exit on keyboard interrupt
def handle_exit():
    print("\nScript terminated by user.")
    sys.exit()


def main(): 
    try:
        # Wait for user to start
        wait_for_start()

        # Define the range of numbers: 
        for number in range(311, 660, 10):
            print(f"\nStarting automation for number: {number}")
            time.sleep(1)

            # Step 1: Click "Search"
            move_and_click(1046, 300)
            time.sleep(4)

            # Step 2: Type number
            move_and_click(192, 317)
            time.sleep(.2)
            type_number(number)
            pyautogui.press('enter')
            time.sleep(5)


            # Step 3: Repeat 10 times
            for i in range(10):
                print(f"  Iteration {i+1} of 10")

                # Click "Select All"
                move_and_click(160, 354)

                # Click "Next Page"
                move_and_click(222, 320)
                time.sleep(2)

            # Step 4: Click "Download"
            move_and_click(853, 320)
            time.sleep(1)

            # Step 5: Click "Continue"
            move_and_click(888, 865)
            time.sleep(8)

            # Step 6: Click "Revise Search"
            move_and_click(998, 282)
            time.sleep(3)
            

            print(f"Automation for number {number} complete!")

            

        print("\nAll automations completed successfully!")

    except KeyboardInterrupt:
        handle_exit()
    except Exception as e:
        print(f"\nAn error occurred: {e}")
        handle_exit()


if __name__ == "__main__":
    main()