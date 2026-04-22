#!/bin/bash
cd "$(dirname "$0")"
echo "Running docx-formatter tests..."
echo ""
echo "=== test_utils.py ==="
python3 test_utils.py -v
echo ""
echo "=== test_fix_quotes.py ==="
python3 test_fix_quotes.py -v
echo ""
echo "=== test_table_borders.py ==="
python3 test_table_borders.py -v
echo ""
echo "=== test_integration.py ==="
python3 test_integration.py -v
echo ""
echo "All tests completed."
