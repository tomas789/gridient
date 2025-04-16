# tests/test_examples.py
import os
import runpy

import pytest

# Get the directory containing this test file
TEST_DIR = os.path.dirname(os.path.abspath(__file__))
# Go up one level to the project root
PROJECT_ROOT = os.path.dirname(TEST_DIR)
EXAMPLES_DIR = os.path.join(PROJECT_ROOT, "examples")

# Map example script names to their expected output files
EXAMPLE_OUTPUT_FILES = {
    "house_power_price.py": "house_power_price_output.xlsx",
    "house_mortgage.py": "house_mortgage_output.xlsx",
}


@pytest.mark.parametrize(
    "example_script",
    EXAMPLE_OUTPUT_FILES.keys(),  # Use keys from the map
)
def test_example_runs_without_error(example_script):
    """
    Test that an example script runs to completion without raising exceptions
    and produces its expected output file.
    """
    script_path = os.path.join(EXAMPLES_DIR, example_script)
    output_filename = EXAMPLE_OUTPUT_FILES[example_script]
    # Assume output file is created in the project root for simplicity
    # Adjust if examples save files elsewhere
    output_file_path = os.path.join(PROJECT_ROOT, output_filename)

    assert os.path.exists(script_path), f"Example script not found: {script_path}"

    # 1. Attempt to remove the output file before running (silent if not found)
    try:
        os.remove(output_file_path)
        print(f"\nRemoved pre-existing output file: {output_filename}")
    except FileNotFoundError:
        print(f"\nNo pre-existing output file to remove: {output_filename}")
        pass  # It's okay if the file doesn't exist yet
    except OSError as e:
        pytest.fail(f"Error removing pre-existing file {output_filename}: {e}")

    print(f"Running example: {example_script}...")
    try:
        # 2. Run the example script
        runpy.run_path(script_path, run_name="__main__")
        print(f"Example {example_script} completed successfully.")
    except Exception as e:
        pytest.fail(f"Example script {example_script} failed with exception: \n{type(e).__name__}: {e}")

    # 3. Check if the output file exists AFTER running the script
    assert os.path.exists(output_file_path), (
        f"Example script {example_script} did not create expected output file: {output_filename}"
    )
    print(f"Verified output file exists: {output_filename}")


# Clean up generated files after tests run (optional - kept for overall cleanup)
@pytest.fixture(scope="session", autouse=True)
def cleanup_generated_files():
    # This code runs before any tests in the session
    yield  # Let the tests run
    # This code runs after all tests in the session
    print("\nFinal cleanup of generated Excel files...")
    files_to_remove = [os.path.join(PROJECT_ROOT, fname) for fname in EXAMPLE_OUTPUT_FILES.values()]

    for f_path in files_to_remove:
        if os.path.exists(f_path):
            try:
                os.remove(f_path)
                print(f" Removed: {os.path.basename(f_path)}")
            except OSError as e:
                print(f" Error removing {os.path.basename(f_path)}: {e}")
