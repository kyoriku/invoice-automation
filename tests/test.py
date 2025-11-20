"""
Test script to verify core functionality of the invoice automation system
"""

import os
import sys
from datetime import datetime, timedelta
from pathlib import Path
from dotenv import load_dotenv

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

dotenv_path = Path(__file__).resolve().parent.parent / ".env"
load_dotenv(dotenv_path)

from src.colours import GREEN, RED, RESET

# Import the main class
try:
    from invoice import BiWeeklyInvoiceAutomator
    print(f"{GREEN}Successfully imported BiWeeklyInvoiceAutomator{RESET}")
except ImportError as e:
    print(f"{RED}Could not import main class: {e}{RESET}")
    sys.exit(1)

def test_environment_variables():
    """Test that all required environment variables are present"""
    print("\nTesting environment variables...")
    
    required_vars = ['GMAIL_USER', 'GMAIL_PASSWORD', 'SPREADSHEET_ID', 'RECIPIENT_EMAIL', 'HOURLY_RATE']
    missing_vars = []
    
    for var in required_vars:
        value = os.getenv(var)
        if value:
            print(f"{GREEN}{var}: {'*' * len(value)} (hidden){RESET}")
        else:
            print(f"{RED}{var}: Not set{RESET}")
            missing_vars.append(var)
    
    if missing_vars:
        print(f"{RED}Missing variables: {missing_vars}{RESET}")
        return False
    else:
        print(f"{GREEN}All environment variables are set{RESET}")
        return True

def test_automator_initialization():
    """Test that the automator class initializes correctly"""
    print("\nTesting automator initialization...")
    
    try:
        automator = BiWeeklyInvoiceAutomator()
        print(f"{GREEN}Automator initialized successfully{RESET}")
        print(f"   Hourly rate: ${automator.HOURLY_RATE}")
        print(f"   HST rate: {automator.HST_RATE * 100}%")
        print(f"   Pay period start: {automator.FIRST_PAY_PERIOD_START}")
        return True, automator
    except Exception as e:
        print(f"{RED}Failed to initialize automator: {e}{RESET}")
        return False, None

def test_time_calculations():
    """Test time conversion and hour calculation functions"""
    print("\nTesting time calculations...")
    
    automator = BiWeeklyInvoiceAutomator()
    all_passed = True
    
    # Test time conversion
    test_cases = [
        ("8:00am", "8:00"),
        ("12:00pm", "12:00"),
        ("3:00pm", "15:00"),
        ("11:59pm", "23:59")
    ]
    
    for input_time, expected in test_cases:
        result = automator.convert_time_to_24h_simple(input_time)
        if result == expected:
            print(f"{GREEN}{input_time} → {result}{RESET}")
        else:
            print(f"{RED}{input_time} → {result} (expected {expected}){RESET}")
            all_passed = False
    
    # Test hour calculation
    test_shifts = [
        ("8:00am", "3:00pm", 7.0),
        ("9:00am", "5:00pm", 8.0),
        ("10:00am", "2:00pm", 4.0)
    ]
    
    for start, end, expected_hours in test_shifts:
        calculated = automator.calculate_shift_hours(start, end)
        if calculated == expected_hours:
            print(f"{GREEN}{start} to {end} = {calculated} hours{RESET}")
        else:
            print(f"{RED}{start} to {end} = {calculated} hours (expected {expected_hours}){RESET}")
            all_passed = False
    
    return all_passed

def test_pay_period_calculation():
    """Test pay period calculation logic"""
    print("\nTesting pay period calculations...")
    
    automator = BiWeeklyInvoiceAutomator()
    
    # Test with a known date
    test_date = datetime(2025, 7, 21)  # Monday July 21, 2025
    pay_period = automator.get_pay_period_for_date(test_date)
    
    if pay_period:
        print(f"{GREEN}Pay period calculated for {test_date.strftime('%Y-%m-%d')}{RESET}")
        print(f"   Start: {pay_period['start'].strftime('%Y-%m-%d')}")
        print(f"   End: {pay_period['end'].strftime('%Y-%m-%d')}")
        print(f"   Period key: {pay_period['period_key']}")
        return True
    else:
        print(f"{RED}Could not calculate pay period for {test_date}{RESET}")
        return False

def test_email_parsing():
    """Test email parsing with sample data"""
    print("\nTesting email parsing...")
    
    automator = BiWeeklyInvoiceAutomator()
    
    # Sample email content
    sample_email = """
    Here's what the week looks like:
    07/21 - 07/27 | Published by Company Name

    Monday 07/21
    8:00am - 3:00pm

    Tuesday 07/22
    8:00am - 3:00pm

    Wednesday 07/23
    8:00am - 3:00pm

    Want some extra hours?
    Here is an available open shift:
    
    Wednesday 07/23
    8:00am - 3:00pm
    """
    
    shifts = automator.parse_homebase_email(sample_email)
    
    if len(shifts) == 3:  # Should exclude the optional shift
        print(f"{GREEN}Parsed {len(shifts)} shifts correctly{RESET}")
        for shift in shifts:
            print(f"   {shift['date']}: {shift['start_time']} - {shift['end_time']}")
        return True
    else:
        print(f"{RED}Expected 3 shifts, got {len(shifts)}{RESET}")
        return False

def test_file_operations():
    """Test file read/write operations"""
    print("\nTesting file operations...")
    
    automator = BiWeeklyInvoiceAutomator()
    
    test_file = "test_shifts_data.json"
    original_file = automator.SHIFTS_DATA_FILE
    
    try:
        automator.SHIFTS_DATA_FILE = test_file
        test_data = {"test": "data", "timestamp": datetime.now().isoformat()}
        
        if automator.save_shifts_data(test_data):
            print(f"{GREEN}Successfully saved test data{RESET}")
            loaded_data = automator.load_shifts_data()
            if loaded_data.get("test") == "data":
                print(f"{GREEN}Successfully loaded test data{RESET}")
                success = True
            else:
                print(f"{RED}Loaded data doesn't match saved data{RESET}")
                success = False
        else:
            print(f"{RED}Failed to save test data{RESET}")
            success = False
            
    except Exception as e:
        print(f"{RED}File operation error: {e}{RESET}")
        success = False
    
    finally:
        automator.SHIFTS_DATA_FILE = original_file
        try:
            if os.path.exists(test_file):
                os.remove(test_file)
        except:
            pass
    
    return success

def run_all_tests():
    """Run all tests and report results"""
    print("Starting functionality tests...\n")
    
    tests = [
        ("Environment Variables", test_environment_variables),
        ("Automator Initialization", test_automator_initialization),
        ("Time Calculations", test_time_calculations),
        ("Pay Period Calculation", test_pay_period_calculation),
        ("Email Parsing", test_email_parsing),
        ("File Operations", test_file_operations)
    ]
    
    results = []
    
    for test_name, test_func in tests:
        try:
            if test_func.__name__ == "test_automator_initialization":
                success, _ = test_func()
            else:
                success = test_func()
            results.append((test_name, success))
        except Exception as e:
            print(f"{RED}{test_name} failed with error: {e}{RESET}")
            results.append((test_name, False))
    
    # Summary
    print("\n" + "="*50)
    print("TEST RESULTS SUMMARY")
    print("="*50)
    
    passed = 0
    total = len(results)
    
    for test_name, success in results:
        status = f"{GREEN}PASS{RESET}" if success else f"{RED}FAIL{RESET}"
        print(f"{status} {test_name}")
        if success:
            passed += 1
    
    print(f"\nTests passed: {passed}/{total}")
    
    if passed == total:
        print(f"{GREEN}All tests passed! Your system is ready to use.{RESET}")
    else:
        print(f"{RED}Some tests failed. Please check the issues above.{RESET}")
    
    return passed == total

if __name__ == "__main__":
    run_all_tests()
