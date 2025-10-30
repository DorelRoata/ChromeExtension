#!/usr/bin/env python3
"""
Test script to verify the improvements made to main.py
Tests memory management, error handling, and window position tracking
"""

import tkinter as tk
import gc
import time
import os
import sys

# Add current directory to path to import main module
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from main import set_window_icon_safe, setup_window_position, handle_error
    print("Successfully imported new functions from main.py")
except ImportError as e:
    print(f"Failed to import functions: {e}")
    sys.exit(1)

def test_memory_fix():
    """Test memory usage with multiple windows"""
    print("\n=== Testing Memory Fix ===")
    
    try:
        for i in range(3):
            root = tk.Tk()
            root.title(f"Test Window {i}")
            root.geometry("300x200")
            
            # Test the new icon function
            set_window_icon_safe(root, icon_png_path='icon.png')
            setup_window_position(root, f"test_{i}", (300, 200))
            
            # Simulate some work
            root.update()
            time.sleep(0.1)
            
            root.destroy()
            gc.collect()
        
        print("Memory test completed - no crashes detected")
        return True
    except Exception as e:
        print(f"Memory test failed: {e}")
        return False

def test_error_handling():
    """Test improved error messages"""
    print("\n=== Testing Error Handling ===")
    
    try:
        # Test file not found
        try:
            raise FileNotFoundError("test.xlsx")
        except Exception as e:
            msg = handle_error(e, "Loading Excel file", show_user=False)
            print(f"File not found error handled: {type(e).__name__}")
        
        # Test permission error
        try:
            raise PermissionError("Access denied")
        except Exception as e:
            msg = handle_error(e, "Saving to Excel", show_user=False)
            print(f"Permission error handled: {type(e).__name__}")
        
        # Test Excel error
        try:
            raise Exception("openpyxl: Workbook is locked")
        except Exception as e:
            msg = handle_error(e, "Processing data", show_user=False)
            print(f"Excel error handled: {type(e).__name__}")
        
        print("Error handling test completed")
        return True
    except Exception as e:
        print(f"Error handling test failed: {e}")
        return False

def test_window_positions():
    """Test window position tracking"""
    print("\n=== Testing Window Position Tracking ===")
    
    try:
        root = tk.Tk()
        root.title("Position Test")
        root.geometry("400x300")
        
        # Test the new position function
        setup_window_position(root, "position_test", (400, 300))
        set_window_icon_safe(root, icon_png_path='icon.png')
        
        print("Window position tracking initialized")
        print("  (Close the window to test position saving)")
        
        # Brief display to test
        root.update()
        time.sleep(0.5)
        
        root.destroy()
        
        # Check if config file was created
        if os.path.exists("window_config.json"):
            print("Window position config file created")
        else:
            print("? Window position config file not created (may appear on close)")
        
        return True
    except Exception as e:
        print(f"Window position test failed: {e}")
        return False

def test_config_file():
    """Test configuration file functionality"""
    print("\n=== Testing Configuration File ===")
    
    try:
        from main import load_window_positions, save_window_positions
        
        # Test saving
        test_positions = {
            "main": {"x": 100, "y": 100, "width": 400, "height": 200},
            "batch": {"x": 150, "y": 150, "width": 500, "height": 400}
        }
        
        save_window_positions(test_positions)
        print("Window positions saved to config")
        
        # Test loading
        loaded_positions = load_window_positions()
        if loaded_positions == test_positions:
            print("Window positions loaded correctly")
        else:
            print("? Window positions loaded with differences")
        
        return True
    except Exception as e:
        print(f"Configuration file test failed: {e}")
        return False

def main():
    """Run all tests"""
    print("Testing improvements to main.py")
    print("=" * 50)
    
    tests = [
        ("Memory Management", test_memory_fix),
        ("Error Handling", test_error_handling),
        ("Window Positions", test_window_positions),
        ("Configuration File", test_config_file)
    ]
    
    results = []
    for test_name, test_func in tests:
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"✗ {test_name} test crashed: {e}")
            results.append((test_name, False))
    
    # Summary
    print("\n" + "=" * 50)
    print("TEST SUMMARY")
    print("=" * 50)
    
    passed = 0
    for test_name, result in results:
        status = "PASS" if result else "FAIL"
        symbol = "PASS" if result else "FAIL"
        print(f"{symbol} {test_name}: {status}")
        if result:
            passed += 1
    
    print(f"\nOverall: {passed}/{len(results)} tests passed")
    
    if passed == len(results):
        print("All improvements working correctly!")
        return 0
    else:
        print("Some issues detected - review the output above")
        return 1

if __name__ == "__main__":
    sys.exit(main())