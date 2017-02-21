package main

import (
	"bufio"
	"fmt"
	"math"
	"os"
	"strings"

	//"pflags"
	"github.com/spf13/pflag"

	"github.com/go-ole/go-ole"
	//"github.com/go-ole/go-ole/oleutil"
	//"github.com/noypi/xlsx"
	"github.com/willauld/xlsx"
	//"github.com/xlsx"
)

var store_items = map[string]string{
	"Full Ingredient Sake Kit": "Rice milled to ~60% 10 lbs.\nKoji 40 Oz.\nYeast #9\nLactic Acid 2 fl. Oz.\nYeast Nutrient 1 Oz.\nSpeedy Bentonite 2 Oz.",
	"Sake Ingredient Kit":      "Rice milled to ~60% 10 lbs.\nKoji 40 Oz.\nYeast #9",
	"Rice milled for Sake":     "Medium grain rice\nMilled to ~60% (Ginjo Level)\n10 lbs. bag",
	"Koji":                     "Rice milled to ~60% cultured with koji kin\n40 oz package",
	"Yeast #9":                 "Wyeast 4134 Sake",
	"Lactic Acid 88%":          "Lactic Acid 88%\n2Fl. Oz.",
	"Yeast Nutrient":           "Thiamin, vitamin B complex\n1 Oz.",
	"Speedy Bentonite":         "Bentonite clay wine clarifier\n2 Oz.",
	"Koji-kin":                 "15g Powdered Rice Koji Starter\nEnough to make 2 batches of 2.5 lbs. koji each\nAspergillus oryzae and rice flour\nPrinted Instructions",
	"Yeast #7":                 "White Labs WLP705",
	"Special Ginjo Koji-kin":   "2 x 1g Powdered Akita Konno Special Ginjo Koji Starter\nEach 1g packet makes 3.14 lbs koji (6.28 lbs. total)\nAspergillus oryzae\nPrinted Instructions",
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
	quantity         int
	title            string
	item_price       float32
	total_item_price float32
}
type order struct {
	billing       address
	shipping      address
	items         []item
	total_item    float32
	shipping_cost float32
	total_order   float32
}

func IntMax(a, b int) int {
	if a > b {
		return a
	}
	return b
}

func getQuantity(line string) int {
	var digit int
	quantity := 0
	first_digit := false

	for i := 0; i < len(line); i += 1 {
		if line[i] >= '0' && line[i] <= '9' {
			first_digit = true
			digit = int(line[i]) - int('0')
			quantity = quantity*10 + digit
		}
		if first_digit == true {
			return quantity
		}
	}
	return 0
}

func getTitle(line string) string {

	var start_char int
	var end_char int
	for i := 0; i < len(line); i += 1 {
		if line[i] >= 'A' && line[i] <= 'Z' || line[i] >= 'a' && line[i] <= 'z' {
			start_char = i
			break
		}
	}
	for i := 0; i < len(line); i += 1 {
		if line[i] == '-' && line[i+1] == '-' {
			end_char = i - 1
			break
		}
		if line[i] == '$' {
			end_char = i - 1
			break
		}
	}
	for j := end_char; ; j -= 1 {
		if line[j] != ' ' && line[j] != '\t' {
			end_char = j + 1
			break
		}
	}
	return line[start_char:end_char]
}

func getPrice(line string) float32 {

	var start_char int
	var price float32
	past_point := 0
	found_point := false
	price = 0.0

	for i := 0; i < len(line); i += 1 {
		if line[i] == '$' {
			start_char = i + 1
			break
		}
	}
	for j := start_char; j < len(line); j += 1 {

		if line[j] >= '0' && line[j] <= '9' {
			price = price*10.0 + float32(int(line[j])-int('0'))
			if found_point == true {
				past_point += 1
			}
		} else if line[j] == '.' {
			found_point = true
		}
	}
	price = price / float32(math.Pow(10.0, float64(past_point)))
	return price
}

func output_spreadsheet(purchase order, xlsxPath string, save bool) {

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

	start_row := 19
	for i, v := range purchase.items {
		row := start_row + i

		loc := fmt.Sprintf("b%d", row)
		cell = sheet1.Range(loc)
		cell.PutValue(v.title)

		loc = fmt.Sprintf("c%d", row)
		cell = sheet1.Range(loc)
		cell.PutValue(store_items[v.title])

		loc = fmt.Sprintf("e%d", row)
		cell = sheet1.Range(loc)
		cell.PutValue(v.quantity)

		loc = fmt.Sprintf("f%d", row)
		sheet1.Range(loc).PutValue(v.item_price) //&&&&&&&&&
		//cell = sheet1.Range(loc)
		//cell.PutValue(v.item_price)

		nl_count := strings.Count(v.title, "\n")
		nl_count = IntMax(nl_count, strings.Count(store_items[v.title], "\n"))
		row_range := fmt.Sprintf("%d:%d", row, row)
		sheet1.Range(row_range).PutRowHeight(12.4 * float64(nl_count+1))
	}
	loc := fmt.Sprintf("g%d", 28)

	cell = sheet1.Range(loc)
	cell.PutValue(purchase.shipping_cost)

	if save == true {
		save_to := "C:\\home\\auld\\personal\\Sake'\\PackingSlips\\email\\" +
			billing.firstName + ".xlsx"
		fmt.Printf("Saveing Excel to: %s\n", save_to)
		workbook.Save(save_to)
	} else {
		sheet1.PrintOut(1, 1, 1) // params: fromPage, toPage, copies
	}
	//os.Exit(1)

	//filepath2 := "c:\\temp\\a.xlsx"
	//workbook.Save(filepath2)
}

func print_purchase_record(purchase order, i int) {
	b := purchase.billing
	s := purchase.shipping
	fmt.Println("")
	fmt.Printf("%d: Shipping:                                            Billing:\n", i)
	fmt.Printf("   =========                                            ========\n")
	s1 := s.firstName + " " + s.lastName
	s2 := b.firstName + " " + b.lastName

	fmt.Printf("%d: %-35s                  %-35s\n", i, s1, s2)
	fmt.Printf("%d: %-35s                  %-35s\n", i, s.street, b.street)
	// cancat first and last and then print in fixed space
	l1 := s.city + ", " + s.state + " " + s.zipCode
	l2 := b.city + ", " + b.state + " " + b.zipCode
	fmt.Printf("%d: %-35s                  %-35s\n", i, l1, l2)
	fmt.Printf("%d: %-35s                  %-35s\n", i, s.country, b.country)
	fmt.Printf("%d: %-35s                  %-35s\n", i, b.phoneNumber, b.email)
	fmt.Println("")
	//fmt.Printf("Cap is %d, Len is %d\n", cap(purchase.items), len(purchase.items))
	for i, v := range purchase.items {
		fmt.Printf("%d - %45s $%6.2f      x%d      $%6.2f\n",
			i, v.title, v.item_price, v.quantity, v.total_item_price)
	}
	fmt.Printf("%78s\n", "=========")
	fmt.Printf("%70s $%6.2f\n", "Item Total", purchase.total_item)
	fmt.Printf("%70s $%6.2f\n", "Shipping", purchase.shipping_cost)
	fmt.Printf("%70s $%6.2f\n", "Order Total", purchase.total_order)
	fmt.Printf("---------------------------\n")
}

func main() {

	//fpath := "C:\\home\\auld\\goDev\\src\\packingSlips\\packingSlipTemplate.xlsx"
	default_xlsx := "packingSlipTemplate.xlsx"
	xlsxPtr := pflag.String("xlsx", default_xlsx, "xlsx template file")
	pathPtr := pflag.String("input", "orders.csv", "input customer file in csv format")
	listItPtr := pflag.Bool("listIt", false, "list the store items")

	pflag.Parse()
	//fmt.Println("input:", *pathPtr)
	//fmt.Println("tail:", pflag.Args())

	if *listItPtr == true {
		for k, v := range store_items {
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
	fmt.Printf("default[%s], xlsxPtr[%s]\n", default_xlsx, *xlsxPtr)
	if default_xlsx == *xlsxPtr {
		cwd, _ := os.Getwd()
		j := cwd + "\\" + *xlsxPtr
		xlsxPtr = &j
	}

	scanner := bufio.NewScanner(f)

	var addr *address
	var parts []string
	i := 0
	total_items := 0
	for scanner.Scan() {

		line := scanner.Text()
		if strings.Contains(line, "First Name:") {

			for j := 0; j < 2; j += 1 {
				if j == 0 {
					addr = &orders[i].billing
				} else {
					addr = &orders[i].shipping
				}
				for exit_loop := false; exit_loop == false; {
					parts = strings.Split(line, ": ")
					scanner.Scan()
					line = scanner.Text()

					switch parts[0] {
					case "First Name":
						addr.firstName = parts[1]
					case "Last Name":
						addr.lastName = parts[1]
					case "Address":
						addr.street = parts[1]
					case "City":
						addr.city = parts[1]
					case "Country":
						addr.country = parts[1]
					case "Country (From above Shipping Calc.)":
						addr.country = parts[1]
					case "Billing State":
						addr.state = parts[1]
					case "State":
						addr.state = parts[1]
					case "Delivery State":
						addr.state = parts[1]
					case "Postal Code":
						addr.zipCode = parts[1]
					case "Email":
						addr.email = parts[1]
						exit_loop = true
					case "Phone":
						addr.phoneNumber = parts[1]
						exit_loop = true
					default:
						from := "Billing"
						if j > 0 {
							from = "Shipping"
						}
						fmt.Printf("****missed: \"%s\" from %s\n", parts[0], from)
						exit_loop = true
					}
				}
			}

			//Collect the purchased items here
			orders[i].items = make([]item, 0) //item_array[total_items:0]
			running_total := float32(0)
			for exit_loop := false; exit_loop == false; {
				scanner.Scan()
				line := scanner.Text()

				if strings.Contains(line, "Total:") {
					exit_loop = true
					orders[i].total_order = getPrice(line)
				} else if strings.Contains(line, "-") {

					quant := getQuantity(line)
					title := getTitle(line)
					tip := getPrice(line)
					price := tip / float32(quant)
					a := item{
						quantity:         quant,
						title:            title,
						item_price:       price,
						total_item_price: tip,
					}
					running_total = running_total + tip

					orders[i].items = append(orders[i].items, a)
					total_items += 1
				}
			}

			orders[i].total_item = running_total
			orders[i].shipping_cost = orders[i].total_order - running_total

			print_purchase_record(orders[i], i+1)

			var input string
			fmt.Printf("Print %s's packing slip? (Y/N/S)\n", orders[i].billing.firstName)
			fmt.Scanf("%s\n", &input)
			for {
				if input == "S" || input == "s" {
					output_spreadsheet(orders[i], *xlsxPtr, true)
					break
				} else if input == "Y" || input == "y" {
					output_spreadsheet(orders[i], *xlsxPtr, false)
					break
				} else if input == "N" || input == "n" {
					break
				} else {
					fmt.Printf("Need a Y, N or S\n")
					fmt.Scanf("%s\n", &input)
				}
			}
			i++
		}
	}
}
