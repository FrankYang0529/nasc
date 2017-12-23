import copy
import six


def move_slide(presentation, old_index, new_index):
    # copy from: https://github.com/scanny/python-pptx/issues/68#issuecomment-129491554

    xml_slides = presentation.slides._sldIdLst  # pylint: disable=W0212
    slides = list(xml_slides)
    xml_slides.remove(slides[old_index])
    xml_slides.insert(new_index, slides[old_index])


def _get_blank_slide_layout(presentation):
    # copy from: https://github.com/scanny/python-pptx/issues/132#issuecomment-346699019
    layout_items_count = [len(layout.placeholders) for layout in presentation.slide_layouts]
    min_items = min(layout_items_count)
    blank_layout_id = layout_items_count.index(min_items)

    return presentation.slide_layouts[blank_layout_id]


def duplicate_slide(presentation, index):
    # copy from: https://github.com/scanny/python-pptx/issues/132#issuecomment-346699019
    """Duplicate the slide with the given index in presentation.

    Adds slide to the end of the presentation"""
    source = presentation.slides[index]

    blank_slide_layout = _get_blank_slide_layout(presentation)
    dest = presentation.slides.add_slide(blank_slide_layout)

    for shp in source.shapes:
        el = shp.element
        newel = copy.deepcopy(el)
        dest.shapes._spTree.insert_element_before(newel, 'p:extLst')

    for key, value in six.iteritems(source.part.rels):
        # Make sure we don't copy a notesSlide relation as that won't exist
        if not "notesSlide" in value.reltype:
            dest.part.rels.add_relationship(value.reltype, value._target, value.rId)

    return dest
