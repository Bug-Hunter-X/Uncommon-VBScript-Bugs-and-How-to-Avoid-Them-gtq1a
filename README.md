# Uncommon VBScript Bugs and Solutions

This repository demonstrates some less-obvious bugs that can arise in VBScript programming and shows ways to mitigate them.  VBScript's weak typing and reliance on COM objects can make it prone to errors if not handled carefully.

## Bugs Covered

* **Late Binding Issues:**  The challenges of using `CreateObject` without proper error handling.
* **Type Mismatches:**  Subtle errors arising from VBScript's weak typing.
* **Implicit Type Coercion:** Unexpected behavior from VBScript's automatic type conversions.
* **Unhandled Exceptions:** The dangers of not implementing proper error handling.
* **Incorrect Object References:** Memory leaks and resource exhaustion due to not releasing COM objects.

## Solutions

The `bugSolution.vbs` file provides improved code with better error handling, explicit type checking (where possible), and proper object management.