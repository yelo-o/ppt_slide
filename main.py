import aspose.slides as slides

# Load Presentation

ppt = slides.Presentation("test.pptx")

# Loop through slides
for index in range(ppt.slides.length):

    # Create a new empty presentation
    with slides.Presentation() as newPpt:

        # Remove default slide
        newPpt.slides[0].remove()

        # Add slide to presentation
        newPpt.slides.add_clone(ppt.slides[index])

        # Save presentation
        newPpt.save("slide_{i}.pptx").format(i = index), slides.export.SaveFormat.PPTX)