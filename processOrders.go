package main

import (
	"bufio"
	"fmt"
	"math"
	"os"
	//"strconv"
	"strings"

	//"pflags"
	ole "github.com/go-ole/go-ole"
	"github.com/spf13/pflag"

	"github.com/noypi/xlsx"
	//"github.com/willauld/xlsx"
	//"github.com/xlsx"
)

var version = struct {
	major int
	minor int
}{2, 2}
var storeItems = map[string]string{
	"Full Ingredient Sake Kit": "Rice milled to ~60% 10 lbs.\nKoji 40 Oz.\nYeast #9\nLactic Acid 1 fl. Oz.\nYeast Nutrient 1 Oz.\nBentonite 1 Oz.\nPotassium Chloride 1 Oz.\nMagnesium Sulfate (AKA Epson salt) 1 Oz.",
	"Sake Ingredient Kit":      "Rice milled to ~60% 10 lbs.\nKoji 40 Oz.\nYeast #9",
	"Rice milled for Sake":     "Medium grain rice\nMilled to ~60% (Ginjo Level)\n10 lbs. bag",
	"Koji":                     "Rice milled to ~60% cultured with koji kin\n40 oz package",
	"Yeast #9":                 "Wyeast 4134 Sake",
	"Lactic Acid 88%":          "Lactic Acid 88%\n1Fl. Oz.",
	"Yeast Nutrient":           "Thiamin, vitamin B complex\n1 Oz.",
	"Bentonite":                "Bentonite clay wine clarifier\n1 Oz.",
	"Koji-kin":                 "15g Powdered Rice Koji Starter\nEnough to make 2 batches of 2.5 lbs. koji each\nAspergillus oryzae and rice flour\nPrinted Instructions",
	"Yeast #7":                 "White Labs WLP705",
	"Special Ginjo Koji-kin":   "2 x 1g Powdered Akita Konno Special Ginjo Koji Starter\nEach 1g packet makes 3.14 lbs koji (6.28 lbs. total)\nAspergillus oryzae\nPrinted Instructions",
	"Potassium Chloride":       "Granulated, 1 oz.",
	"Magnesium Sulfate":        "Granulated, 1 oz.\nAKA Epson salt",
}

type address struct {
	firstName   string
	lastName    string
	street      string
	city        string
	state       string
	country     string
	zipCode     string
	email       string
	phoneNumber string
}
type item struct {
	quantity       int
	title          string
	itemPrice      float32
	totalItemPrice float32
}
type order struct {
	billing      address
	shipping     address
	items        []item
	totalItem    float32
	shippingCost float32
	totalOrder   float32
}

func intMax(a, b int) int {
	if a > b {
		return a
	}
	return b
}

// getQuantity gets the last integer before '$'
func getQuantity(line string) int {
	var digit int
	quantity := 0
	tmpVal := 0
	//firstDigit := false
	lastCharADigit := false

	for i := 0; i < len(line); i++ {
		if line[i] >= '0' && line[i] <= '9' {
			//firstDigit = true
			digit = int(line[i]) - int('0')
			tmpVal = tmpVal*10 + digit
			lastCharADigit = true
		} else if lastCharADigit == true {
			quantity = tmpVal
			tmpVal = 0
			lastCharADigit = false
		} else if line[i] == '$' {
			return quantity
		}
	}
	return 0
}

// getTitle gets the first substring starting with [AZaz] to either a '$' or '--' sequence
func getTitle(line string) string {

	startChar := 0
	endChar := 0
	for i := 0; i < len(line); i++ {
		if line[i] >= 'A' && line[i] <= 'Z' || line[i] >= 'a' && line[i] <= 'z' {
			startChar = i
			break
		}
	}
	for i := 0; i < len(line); i++ {
		if line[i] == '-' && line[i+1] == '-' {
			endChar = i - 1
			break
		}
		if line[i] == '$' {
			endChar = i - 1
			break
		}
	}
	if endChar == 0 {
		endChar = len(line) - 1
	}
	for j := endChar; j >= 0; j-- {
		if line[j] != ' ' && line[j] != '\t' {
			endChar = j + 1
			break
		}
	}
	return line[startChar:endChar]
}

// getPrice returns the first float number passed the '$'
func getPrice(line string) float32 {

	var startChar int
	var price float32
	pastPoint := 0
	foundPoint := false
	price = 0.0

	for i := 0; i < len(line); i++ {
		if line[i] == '$' {
			startChar = i + 1
			break
		}
	}
	for j := startChar; j < len(line); j++ {

		if line[j] >= '0' && line[j] <= '9' {
			price = price*10.0 + float32(int(line[j])-int('0'))
			if foundPoint == true {
				pastPoint++
			}
		} else if line[j] == '.' {
			foundPoint = true
		}
	}
	price = price / float32(math.Pow(10.0, float64(pastPoint)))
	return price
}

// outputSpreadsheet updates the template spreadsheet and prints it
func outputSpreadsheet(purchase order, xlsxPath string, save bool) {

	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	var err error
	defer func() {
		if nil != err {
			fmt.Println("err=", err)
		}
	}()

	excel, err := xlsx.CreateObject()
	if nil != err {
		fmt.Println(err)
		return
	}
	defer excel.Release()

	//xlsxPath := "C:\\home\\auld\\goDev\\src\\packingSlips\\packingSlipTemplate.xlsx"

	workbooks := excel.Workbooks()
	//workbook := workbooks.Create()
	workbook := workbooks.Open(xlsxPath)
	defer workbook.Close()

	sheet1 := workbook.Worksheets(1)

	//cell := sheet1.Range("a1")
	//a1Val := cell.ToString()
	//fmt.Println("a1Val: ", a1Val)
	//cell.PutValue("adrian guwapo")

	shipping := &purchase.shipping
	billing := &purchase.billing

	cell := sheet1.Range("c12")
	cell.PutValue(shipping.firstName + " " + shipping.lastName)

	cell = sheet1.Range("c13")
	cell.PutValue(shipping.street)

	cell = sheet1.Range("c14")
	cell.PutValue(shipping.city + ", " + shipping.state + " " + shipping.zipCode)

	cell = sheet1.Range("c15")
	cell.PutValue(shipping.country)

	cell = sheet1.Range("f12")
	cell.PutValue(billing.firstName + " " + billing.lastName)

	cell = sheet1.Range("f13")
	cell.PutValue(billing.street)

	cell = sheet1.Range("f14")
	cell.PutValue(billing.city + ", " + billing.state + " " + billing.zipCode)

	cell = sheet1.Range("f15")
	cell.PutValue(billing.country)

	cell = sheet1.Range("f16")
	cell.PutValue(billing.email)

	startRow := 19
	for i, v := range purchase.items {
		row := startRow + i

		loc := fmt.Sprintf("b%d", row)
		cell = sheet1.Range(loc)
		cell.PutValue(v.title)

		loc = fmt.Sprintf("c%d", row)
		cell = sheet1.Range(loc)
		cell.PutValue(storeItems[v.title])

		loc = fmt.Sprintf("e%d", row)
		cell = sheet1.Range(loc)
		cell.PutValue(v.quantity)

		loc = fmt.Sprintf("f%d", row)
		sheet1.Range(loc).PutValue(v.itemPrice) //&&&&&&&&&
		//cell = sheet1.Range(loc)
		//cell.PutValue(v.itemPrice)

		nlCount := strings.Count(v.title, "\n")
		nlCount = intMax(nlCount, strings.Count(storeItems[v.title], "\n"))
		rowRange := fmt.Sprintf("%d:%d", row, row)
		sheet1.Range(rowRange).PutRowHeight(12.4 * float64(nlCount+1))
	}
	loc := fmt.Sprintf("g%d", 28)

	cell = sheet1.Range(loc)
	cell.PutValue(purchase.shippingCost)

	if save == true {
		saveTo := "C:\\home\\auld\\personal\\Sake'\\PackingSlips\\email\\" +
			billing.firstName + ".xlsx"
		fmt.Printf("Saveing Excel to: %s\n", saveTo)
		workbook.Save(saveTo)
	} else {
		sheet1.PrintOut(1, 1, 1) // params: fromPage, toPage, copies
	}
	//os.Exit(1)

	//filepath2 := "c:\\temp\\a.xlsx"
	//workbook.Save(filepath2)
}

// printPurchaseRecord displays the purchase information on the terminal
func printPurchaseRecord(purchase order, i int) {
	b := purchase.billing
	s := purchase.shipping
	fmt.Println("")
	fmt.Printf("%d: Shipping:                            Billing:\n", i)
	fmt.Printf("   =========                            ========\n")
	s1 := s.firstName + " " + s.lastName
	s2 := b.firstName + " " + b.lastName

	fmt.Printf("%d: %-35s  %-35s\n", i, s1, s2)
	fmt.Printf("%d: %-35s  %-35s\n", i, s.street, b.street)
	// cancat first and last and then print in fixed space
	l1 := s.city + ", " + s.state + " " + s.zipCode
	l2 := b.city + ", " + b.state + " " + b.zipCode
	fmt.Printf("%d: %-35s  %-35s\n", i, l1, l2)
	fmt.Printf("%d: %-35s  %-35s\n", i, s.country, b.country)
	fmt.Printf("%d: %-35s  %-35s\n", i, b.phoneNumber, b.email)
	fmt.Println("")
	//fmt.Printf("Cap is %d, Len is %d\n", cap(purchase.items), len(purchase.items))
	for i, v := range purchase.items {
		fmt.Printf("%d - %45s $%6.2f      x%d      $%6.2f\n",
			i, v.title, v.itemPrice, v.quantity, v.totalItemPrice)
	}
	fmt.Printf("%78s\n", "=========")
	fmt.Printf("%70s $%6.2f\n", "Item Total", purchase.totalItem)
	fmt.Printf("%70s $%6.2f\n", "Shipping", purchase.shippingCost)
	fmt.Printf("%70s $%6.2f\n", "Order Total", purchase.totalOrder)
	fmt.Printf("---------------------------\n")
}
func empty(s string) bool {
	return len(strings.TrimSpace(s)) == 0
}

// nonEmptyLine
func nonEmptyLine(scanner *bufio.Scanner) (str string) {
	// ignore anything currently in the buffer
	//str = strings.TrimSpace(scanner.Text())
	for scanner.Scan() {
		str = strings.TrimSpace(scanner.Text())
		if str != "" {
			break
		}
	}
	//fmt.Printf("nonEmptyLine: %s\n", str)
	return str
}

// Pattern is a type code
type Pattern int

const (
	noPattern Pattern = iota
	singleWordPattern
	singletonZipPattern
	zipLinePattern
	phonePattern
	emailPattern
)

func patternMatch(str string) Pattern {
	if strings.Contains(str, "@") {
		return emailPattern
	}
	str2 := strings.TrimSpace(str)
	slen := 0
	i := 0
	if str2[i] == '+' {
		i++
		slen++
	}
	for ; i < len(str2); i++ {
		if '0' <= str2[i] && str2[i] <= '9' ||
			str2[i] == ' ' || str2[i] == '-' {
			slen++
		}
	}
	if len(str2) == slen {
		if slen > 5 { // skip zip code on its own line
			return phonePattern
		}
		return singletonZipPattern
	}
	if strings.Contains(str2, ",") {
		s := strings.Split(str2, ",")
		//city := s[0]
		s2 := strings.Split(strings.TrimSpace(s[1]), " ")
		if len(s2) < 2 {
			// no following zip code
			return noPattern
		}
		code := s2[1]
		slen := 0
		for i := 0; i < len(code); i++ {
			if '0' <= code[i] && code[i] <= '9' {
				slen++
			}
		}
		if len(code) == slen { // should check more on this patter
			return zipLinePattern
		}
	}
	word := str2
	slen = 0
	for i := 0; i < len(word); i++ {
		if 'a' <= word[i] && word[i] <= 'z' ||
			'A' <= word[i] && word[i] <= 'Z' {
			slen++
		}
	}
	if len(word) == slen {
		return singleWordPattern
	}
	// first last names // i == 1
	// city, state zip // zip line pattern i == n
	// country i == n+1 // singleWord pattern
	// phone // phone pattern
	// email // email pattern
	return noPattern
}
func explodeZipLine(l string) (city, state, code string) {
	if strings.Contains(l, ",") {
		s := strings.Split(l, ", ")
		city := s[0]
		s2 := strings.Split(strings.TrimSpace(s[1]), " ")
		state := s2[0]
		code := s2[1]
		slen := 0
		for i := 0; i < len(code); i++ {
			if '0' <= code[i] && code[i] <= '9' {
				slen++
			}
		}
		if len(code) == slen { // should check more on this patter
			return city, state, code
		}
	}
	return "", "", ""
}

func readCustomerData(ordersp *[]order, i int, scanner *bufio.Scanner) {
	var addr *address
	var parts []string
	var toString *string

	line := nonEmptyLine(scanner)

	dataSections := 2
	// for loop to record customer data
	for j := 0; j < dataSections; j++ {
		if strings.Contains(line, "Billing address") {
			//fmt.Printf("J:%d, Billing\n", j)
			addr = &(*ordersp)[i].billing
		} else if strings.Contains(line, "Shipping address") {
			//fmt.Printf("J:%d, Shipping\n", j)
			addr = &(*ordersp)[i].shipping
		} else if strings.Contains(line, "Note:") {
			// print out the note to stdout
			fmt.Printf("***%s\n***\n", line)
			line = nonEmptyLine(scanner)
			j--
			continue
		} else {
			fmt.Printf("how'd I get here: %s\n", line)
			continue
		}
		i := 0
		for exitLoop := false; exitLoop == false && i < 15; {
			i++
			line = nonEmptyLine(scanner)
			// want to do:
			// name // i == 1
			// addr1// i == 2
			// addr2// i == 3 // ?
			// city, state zip // zip line pattern i == n
			// country i == n+1
			// phone // phone pattern
			// email // email pattern
			pattern := patternMatch(line)
			switch pattern {
			case singleWordPattern:
				// "State/Province":
				toString = &addr.state
				if *toString == "" {
					*toString = line
				} else {
					// "country":
					toString = &addr.country
					if *toString != "" {
						fmt.Printf("Warning: Country being set twice!")
					}
					*toString = line
				}
			case singletonZipPattern:
				szip := line
				// "Postal Code":
				toString = &addr.zipCode
				if *toString != "" {
					fmt.Printf("Warning: zip code being set twice!")
				}
				*toString = szip
			case zipLinePattern:
				city, state, zip := explodeZipLine(line)
				// "City":
				toString = &addr.city
				*toString = city
				// "State/Province":
				toString = &addr.state
				if *toString != "" {
					fmt.Printf("Warning: State being set twice!")
				}
				*toString = state
				// "Postal Code":
				toString = &addr.zipCode
				if *toString != "" {
					fmt.Printf("Warning: zip code being set twice!")
				}
				*toString = zip
			case phonePattern:
				toString = &addr.phoneNumber
				*toString = line
			case emailPattern:
				toString = &addr.email
				*toString = line
			case noPattern:
				if strings.Contains(line, "Shipping address") ||
					strings.Contains(line, "Home Brew Sake") ||
					strings.Contains(line, "Congratulations") {
					// exit billing address loop
					exitLoop = true
				} else if i == 1 {
					// Name, break into first and last
					s := strings.Split(line, " ")
					// first name
					toString = &addr.firstName
					*toString = s[0]
					// last name
					toString = &addr.lastName
					*toString = s[len(s)-1]
				} else if i > 1 {
					// Address
					toString = &addr.street
					if *toString == "" {
						*toString = line
					} else {
						*toString += ", " + line
						fmt.Printf("\n****updated/combined: \"%s\" \n\n", *toString)
					}
				}

				/*
					case "Country":
						toString = &addr.country
						*toString = parts[1]
				*/

			default:
				if !strings.Contains(parts[0], ":") {
					// Idea here is to append this line to the last input
					// For example: street address appended with apartment number
					if toString == nil {
						//log.Fatal("readCustomerData()::toString is nil!!!")
						//tmp := ""
						//toString = &tmp
						// may be useful
						// However, if toSting is nil, we haven't got to the data block yet
						// so just ignore it (FYI there is a "Purchase # xxxx" here that
						// may be useful)
					} else {
						if parts[0] != "" {
							*toString += ", " + parts[0] // note index 0 no 1
							fmt.Printf("\n****updated/combined: \"%s\" \n\n", *toString)
						}
					}
				} else {
					fmt.Printf("\n****missed: \"%s\" \n\n", parts[0])
				}
			}
		}
	} // end for loop to record customer data
}

func collectItemsPurchaseReport(orderp *order, scanner *bufio.Scanner) {
	totalItems := 0

	//Collect the purchased items here (purchase report)
	orderp.items = make([]item, 0) //item_array[totalItems:0]
	runningTotal := float32(0)
	// for loop to load purchase report data
	for exitLoop := false; exitLoop == false; {
		scanner.Scan()
		line := scanner.Text()

		if strings.Contains(line, "Total:") {
			exitLoop = true
			orderp.totalOrder = getPrice(line)
		} else if strings.Contains(line, "Subtotal:") {
			// do nothing for now
		} else if strings.Contains(line, "Shipping:") {
			// do nothing for now
		} else if strings.Contains(line, "Payment method:") {
			// do nothing for now
		} else { // for woo Assume these lines are items

			quant := getQuantity(line)
			title := getTitle(line)
			tip := getPrice(line)
			price := tip / float32(quant)
			a := item{
				quantity:       quant,
				title:          title,
				itemPrice:      price,
				totalItemPrice: tip,
			}
			runningTotal = runningTotal + tip

			orderp.items = append(orderp.items, a)
			totalItems++
		}
	} //end load purchase report data

	orderp.totalItem = runningTotal
	orderp.shippingCost = orderp.totalOrder - runningTotal
}

func createPackingSlip(curOrder order, xlsx string) {
	var input string
	fmt.Printf("Print %s's packing slip? (Y/N/S)\n", curOrder.billing.firstName)
	fmt.Scanf("%s\n", &input)
	for {
		if input == "S" || input == "s" {
			outputSpreadsheet(curOrder, xlsx, true)
			break
		} else if input == "Y" || input == "y" {
			outputSpreadsheet(curOrder, xlsx, false)
			break
		} else if input == "N" || input == "n" {
			break
		} else {
			fmt.Printf("Need a Y, N or S\n")
			fmt.Scanf("%s\n", &input)
		}
	}
}

// main processes customer input csv file, displaying the orders and printer a packing list
func main() {

	//fpath := "C:\\home\\auld\\goDev\\src\\packingSlips\\packingSlipTemplate.xlsx"
	defaultXlsx := "packingSlipTemplate.xlsx"
	versionPtr := pflag.Bool("version", false, "program version")
	xlsxPtr := pflag.String("xlsx", defaultXlsx, "xlsx template file")
	pathPtr := pflag.String("input", "orders.csv", "input customer file in csv format")
	listItPtr := pflag.Bool("listIt", false, "list the store items")

	pflag.Parse()
	//fmt.Println("input:", *pathPtr)
	//fmt.Println("tail:", pflag.Args())

	if *versionPtr == true {
		fmt.Printf("\t Version %d.%d", version.major, version.minor)
		os.Exit(0)
	}
	if *listItPtr == true {
		for k, v := range storeItems {
			fmt.Printf("key[%s] value[%s]\n", k, v)
		}
		os.Exit(0)
	}

	var orders = make([]order, 100)

	//
	// Open Customer data file
	//
	f, err := os.Open(*pathPtr)
	if err != nil {
		fmt.Printf("could not open CVS file: %s\n", *pathPtr)
		os.Exit(1)
		//panic(e)
	}
	//fmt.Printf("default[%s], xlsxPtr[%s]\n", defaultXlsx, *xlsxPtr)
	if defaultXlsx == *xlsxPtr {
		cwd, _ := os.Getwd()
		j := cwd + "\\" + *xlsxPtr
		xlsxPtr = &j
	}

	scanner := bufio.NewScanner(f)
	i := 0
	for scanner.Scan() {
		line := scanner.Text()
		if strings.Contains(line, "Product	 Quantity	 Price") {
			//fmt.Printf("found product heading\n")
			collectItemsPurchaseReport(&orders[i], scanner)
			readCustomerData(&orders, i, scanner)
			printPurchaseRecord(orders[i], i+1)
			createPackingSlip(orders[i], *xlsxPtr)
			i++
		}
	}
}
