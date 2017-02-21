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
		r := IntMax(elem.a, elem.b)
		if elem.result != r {
			t.Errorf("unexpected result: intMax(%d,%d) returned %d but expected %d\n", elem.a, elem.b, r, elem.result)
		}
	}
}

func TestGetQuantity(t *testing.T) {

}
