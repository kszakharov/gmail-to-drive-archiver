TEMPLATE=main.template.gs
OUTPUT=main.gs
ENV=.env

build:
	@set -a; \
	source $(ENV); \
	set +a; \
	sed \
		-e "s|__DUPLICATE_MODE__|$$DUPLICATE_MODE|g" \
		-e "s|__FOLDER_ID__|$$FOLDER_ID|g" \
		-e "s|__GRANULARITY__|$$GRANULARITY|g" \
		-e "s|__INITIAL_LAST_RUN__|$$INITIAL_LAST_RUN|g" \
		-e "s|__SEARCH_QUERY__|$$SEARCH_QUERY|g" \
		$(TEMPLATE) > $(OUTPUT)
	@echo "Built $(OUTPUT)"
