package main

import "testing"

func TestIntMax(t *testing.T) {
	tests := []struct {
		a, b, result int
	}{
		{1, 5, 5},
		{3, 2, 3},
		{3, 3, 3},
	}
	for _, elem := range tests {
		r := intMax(elem.a, elem.b)
		if elem.result != r {
			t.Errorf("unexpected result: intMax(%d,%d) returned %d but expected %d\n", elem.a, elem.b, r, elem.result)
		}
	}
}

func TestGetQuantity(t *testing.T) {
	testCases := []struct {
		line  string
		quant int
	}{
		{"- 1 other stuff", 1},
		{" some stuff 5 other stuff", 5},
		{"- 10 other stuff", 10},
		{"- 162 other stuff", 162},
		{"-1 other stuff", 1},
	}
	for _, testCase := range testCases {
		quantity := getQuantity(testCase.line)
		if quantity != testCase.quant {
			t.Errorf("getQuantity(%s) returned %d but expected %d\n", testCase.line, quantity, testCase.quant)
		}
	}
}

func TestGetTitle(t *testing.T) {
	tstCases := []struct {
		line     string
		expected string
	}{
		{"- 56 A title 76 -- not this $ 6.43", "A title 76"},
		{"- A title -- not this $ 6.43", "A title"},
		{"- A title not this $ 6.43", "A title not this"},
		{"- A title ", "A title"},
		{"A title ", "A title"},
	}
	for _, tstCase := range tstCases {
		title := getTitle(tstCase.line)
		if title != tstCase.expected {
			t.Errorf("getTitle(%s) returned [%s] but [%s] was expected\n", tstCase.line, title, tstCase.expected)
		}
	}
}

func TestGetPrice(t *testing.T) {
	tstCases := []struct {
		line     string
		expected float32
	}{
		{"$6.50", 6.50},
		{" - 23 This is the title -- other stuff $ 6.50  ", 6.50},
	}
	for _, tstCase := range tstCases {
		price := getPrice(tstCase.line)
		if price != tstCase.expected {
			t.Errorf("getPrice(%s) returned [%6.2f] expected [%6.2f]\n", tstCase.line, price, tstCase.expected)
		}
	}
}
