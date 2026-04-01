const STATUS_SHAPE_NAME = "SlideStatusMarker";

const STATUSES = {
  toBeUpdated: { label: "To be updated", icon: "🔴" },
  draft: { label: "Draft", icon: "🟠" },
  forReview: { label: "For review", icon: "🔵" },
  completed: { label: "Completed", icon: "🟢" }
};

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    setMessage("This add-in only works in PowerPoint.", true);
    return;
  }

  renderButtons();
  setMessage("Ready. Select a status to apply.");
});

function renderButtons() {
  const singleContainer = document.getElementById("single-slide-actions");
  const allContainer = document.getElementById("all-slide-actions");

  Object.entries(STATUSES).forEach(([key, status]) => {
    singleContainer.appendChild(createActionButton(`${status.icon} ${status.label}`, () => setCurrentSlideStatus(key)));
    allContainer.appendChild(createActionButton(`${status.icon} ${status.label}`, () => setAllSlidesStatus(key)));
  });
}

function createActionButton(text, onClick) {
  const button = document.createElement("button");
  button.textContent = text;
  button.addEventListener("click", onClick);
  return button;
}

async function setCurrentSlideStatus(statusKey) {
  const status = STATUSES[statusKey];
  if (!status) {
    return;
  }

  try {
    await PowerPoint.run(async (context) => {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items/id");
      await context.sync();

      if (selectedSlides.items.length === 0) {
        throw new Error("Please select a slide first.");
      }

      const slide = context.presentation.slides.getItem(selectedSlides.items[0].id);
      await upsertStatusShape(context, slide, status);
      await context.sync();
    });

    setMessage(`Updated current slide to: ${status.label}`);
  } catch (error) {
    setMessage(error.message || "Could not update the current slide.", true);
  }
}

async function setAllSlidesStatus(statusKey) {
  const status = STATUSES[statusKey];
  if (!status) {
    return;
  }

  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();

      for (const slide of slides.items) {
        await upsertStatusShape(context, slide, status);
      }

      await context.sync();
    });

    setMessage(`Updated all slides to: ${status.label}`);
  } catch (error) {
    setMessage(error.message || "Could not update all slides.", true);
  }
}

async function upsertStatusShape(context, slide, status) {
  const shapes = slide.shapes;
  shapes.load("items/name");
  await context.sync();

  const text = `${status.icon} ${status.label}`;
  const existing = shapes.items.find((shape) => shape.name === STATUS_SHAPE_NAME);

  if (existing) {
    existing.textFrame.textRange.text = text;
    return;
  }

  const newShape = shapes.addTextBox(text, {
    left: 18,
    top: 18,
    width: 170,
    height: 28
  });

  newShape.name = STATUS_SHAPE_NAME;
}

function setMessage(message, isError = false) {
  const el = document.getElementById("message");
  el.textContent = message;
  el.classList.toggle("error", isError);
}
