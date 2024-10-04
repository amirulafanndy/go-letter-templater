package main

import (
	"fmt"

	"github.com/unidoc/unioffice/document"
)

// Data struct that holds the data to be injected
type Person struct {
	Name  string
	Date  string
	Email string
}

func main() {
	// Simulated data that could come from a database, form, or API
	person := Person{
		Name:  "John Doe",
		Date:  "2024-10-01",
		Email: "johndoe@example.com",
	}

	// Process the DOCX file
	err := mergeDataIntoDocx(person, "template.docx", "output.docx")
	if err != nil {
		fmt.Println("Error:", err)
	} else {
		fmt.Println("Document saved successfully as output.docx")
	}
}

func mergeDataIntoDocx(person Person, inputPath string, outputPath string) error {
	// Open the DOCX template
	doc, err := document.Open(inputPath)
	if err != nil {
		return fmt.Errorf("could not open docx: %v", err)
	}

	// Loop through paragraphs and runs to replace placeholders
	for _, para := range doc.Paragraphs() {
		for _, run := range para.Runs() {
			text := run.Text()

			// Replace merge fields with actual data
			if text == "{{Name}}" {
				run.ClearContent()
				run.AddText(person.Name)
			}
			if text == "{{Date}}" {
				run.ClearContent()
				run.AddText(person.Date)
			}
			if text == "{{Email}}" {
				run.ClearContent()
				run.AddText(person.Email)
			}
		}
	}

	// Save the updated DOCX file
	err = doc.SaveToFile(outputPath)
	if err != nil {
		return fmt.Errorf("could not save docx: %v", err)
	}
	return nil
}
