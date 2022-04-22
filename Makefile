# Intended for local development only

SHELL := /bin/bash
.DEFAULT_GOAL := help
.PHONY: coverage help install lint push test tox

dir_package="blearn"
dir_tests="tests"


install:  ## Install in local system
	flit install --symlink --extras test

lint:  ## Lint and static-check
	black --check --diff --color $(dir_package) $(dir_tests)
	python -m flake8 $(dir_package) $(dir_tests)
	# python -m pylint $(dir_package)
	python -m mypy $(dir_package) $(dir_tests)

test:  ## Test
	pytest  # configuration defined in pyproject.toml


help: ## Show help message
	@IFS=$$'\n' ; \
	help_lines=(`fgrep -h "##" $(MAKEFILE_LIST) | fgrep -v fgrep | sed -e 's/\\$$//' | sed -e 's/##/:/'`); \
	printf "%s\n\n" "Usage: make [task]"; \
	printf "%-15s %s\n" "Task"            "Help"                          ; \
	printf "%-15s %s\n" "---------------" "-----------------------------" ; \
	for help_line in $${help_lines[@]}; do \
		IFS=$$':' ; \
		help_split=($$help_line) ; \
		help_command=`echo $${help_split[0]} | sed -e 's/^ *//' -e 's/ *$$//'` ; \
		help_info=`echo $${help_split[2]} | sed -e 's/^ *//' -e 's/ *$$//'` ; \
		printf '\033[36m'; \
		printf "%-15s %s" $$help_command ; \
		printf '\033[0m'; \
		printf "%s\n" $$help_info; \
	done
