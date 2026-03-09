CAPABILITY_MANIFEST = """
=== Available PowerPoint Engine Tools ===

SESSION MANAGEMENT:
- pptx_open_presentation(file_path: str) -> {presentation_id, slide_count}
- pptx_create_presentation(template_path?: str) -> {presentation_id}
- pptx_save_presentation(presentation_id, output_path: str) -> {saved_path}
- pptx_close_presentation(presentation_id) -> {closed}
- pptx_get_presentation_state(presentation_id) -> {slide_count, slides: [{index, title, layout}]}

DISCOVERY (read-only):
- pptx_get_layouts(presentation_id) -> {layouts: [{name, placeholder_types}]}
- pptx_get_layout_detail(presentation_id, layout_name) -> {placeholders: [...]}
- pptx_get_masters(presentation_id) -> {masters}
- pptx_get_theme(presentation_id) -> {theme: {colors, fonts}}

SLIDE OPERATIONS:
- pptx_get_slide(presentation_id, slide_index) -> {shapes, placeholders, notes}
- pptx_get_slide_shapes(presentation_id, slide_index) -> lightweight shape list
- pptx_get_slide_text(presentation_id, slide_index) -> all text with formatting
- pptx_add_slide(presentation_id, layout_name, position?) -> {added_slide_index}
- pptx_duplicate_slide(presentation_id, source_index, target_position?) -> {duplicated_slide_index}
- pptx_delete_slide(presentation_id, slide_index) -> {deleted: true}
- pptx_reorder_slides(presentation_id, new_order: int[]) -> {presentation_state}
- pptx_move_slide(presentation_id, from_index, to_index) -> {}
- pptx_set_slide_background(presentation_id, slide_index, color_hex?) -> {}
- pptx_set_slide_notes(presentation_id, slide_index, notes_text) -> {}

TEXT & PLACEHOLDERS:
- pptx_get_placeholders(presentation_id, slide_index) -> [{name, type, left, top, width, height}]
- pptx_set_placeholder_text(presentation_id, slide_index, placeholder_name, text_content, font_name?, font_size_pt?, bold?, italic?, color_hex?) -> {}
- pptx_set_placeholder_rich_text(presentation_id, slide_index, placeholder_name, paragraphs) -> {}
- pptx_get_placeholder_text(presentation_id, slide_index, placeholder_name) -> {text, paragraphs}
- pptx_clear_placeholder(presentation_id, slide_index, placeholder_name) -> {}
- pptx_find_replace_text(presentation_id, find_text, replace_text, case_sensitive?, slide_indices?) -> {}

SHAPES:
- pptx_get_shape_details(presentation_id, slide_index, shape_name|shape_id) -> full shape info
- pptx_add_shape(presentation_id, slide_index, shape_type, left, top, width, height, fill_hex?, text?) -> {}
- pptx_add_text_box(presentation_id, slide_index, left, top, width, height, text_content?) -> {}
- pptx_delete_shape(presentation_id, slide_index, shape_name|shape_id) -> {}
- pptx_set_shape_text(presentation_id, slide_index, shape_name|shape_id, text_content|paragraphs) -> {}
- pptx_set_shape_properties(presentation_id, slide_index, shape_name|shape_id, left?, top?, width?, height?, fill_hex?, line_hex?, name?) -> {}
- pptx_clone_shape(presentation_id, slide_index, shape_name|shape_id, target_slide_index?) -> {}
- pptx_add_image(presentation_id, slide_index, image_path, left?, top?, width?, height?) -> {}
- pptx_copy_shape_between_decks(source_presentation_id, source_slide_index, shape_name|shape_id, target_presentation_id, target_slide_index) -> {}

TABLES:
- pptx_add_table(presentation_id, slide_index, rows, cols, left, top, width, height, data?) -> {}
- pptx_get_table(presentation_id, slide_index, shape_name|shape_id) -> {rows: [[{text}]]}
- pptx_set_table_cell(presentation_id, slide_index, shape_name|shape_id, row, col, text) -> {}
- pptx_set_table_data(presentation_id, slide_index, shape_name|shape_id, data: 2D array) -> {}

CHARTS:
- pptx_add_chart(presentation_id, slide_index, chart_type, categories, series, ...) -> {}
- pptx_get_chart_data(presentation_id, slide_index, shape_name|shape_id) -> {categories, series}
- pptx_update_chart_data(presentation_id, slide_index, shape_name|shape_id, categories?, series) -> {}

=== Parameter conventions ===
- presentation_id: UUID string from pptx_open_presentation or pptx_create_presentation
- slide_index: 1-based integer
- Measurements: "2in", "24pt", "5cm", or raw EMU integer
- Colors: "#RRGGBB" or "RRGGBB"
- shape_name OR shape_id is required for shape operations (not both needed)
- Use $variable_name to reference values produced by earlier plan steps
"""
