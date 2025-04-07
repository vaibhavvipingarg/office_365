.PHONY: setup dev start build certs

setup:
	@echo "Setting up Office 365 Salesforce Add-in..."
	npm install

certs:
	@echo "Installing development certificates..."
	npx office-addin-dev-certs install

dev:
	@echo "Starting development server..."
	npm run dev

start:
	@echo "Starting the add-in..."
	npm start

build:
	@echo "Building for production..."
	npm run build

validate:
	@echo "Validating manifest..."
	npm run validate

all: setup certs

help:
	@echo "Available commands:"
	@echo "  make setup     - Install dependencies"
	@echo "  make certs     - Install development certificates"
	@echo "  make dev       - Start development server"
	@echo "  make start     - Start the add-in"
	@echo "  make build     - Build for production"
	@echo "  make validate  - Validate the manifest"
	@echo "  make all       - Run setup and certs"
	@echo "  make help      - Show this help message" 