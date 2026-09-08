package main

import (
	"log"
	"safety/updates"

	"github.com/joho/godotenv"
)

func init() {

	// Disable log prefixes such as the default timestamp
	log.SetFlags(0)

	// Load the environment file
	if err := godotenv.Load(); err != nil {
		log.Fatalf("Error loading env file: %v", err)
	}
}

func main() {
	d := updates.New()
	if err := d.UpdateReport(); err != nil {
		log.Fatal(err)
	}
}
