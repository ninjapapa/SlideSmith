## Epic 1 Creation of a PowerPoint Interface

The goal is to develop an interface, called PT, that allows for the input of requirements which then generates a PowerPoint deck, possibly even just a screenshot, using GPT to run Python PPTX code.

## Epic 2 Page Content Generation

The second point discusses the need to generate content for individual PowerPoint slides based on descriptions provided in JSON format. This includes the need for a tool called "Harpon High" which seems to be used to make adjustments to the content.

## Epic 3 Storyline Management

The third point emphasizes managing the overall storyline of the presentation, referred to as "the straw man", which involves assembling basic content and a descriptive note on what each slide should entail.

Based on Echo Scribe
```
Audio_STT -> Summary -> Storyline
             ↓
         Simple Memo


Summary -> Storyline

- Gather
  Context & Process
  Audience & Query
  Goal & Choices
  Story outlining

- Flow
  - Confirm on context fields
  - Work on Story as bullets
  - Create page-structure
    + Notes & Process
  - Work on notes
  - Create JSON
```

## Epic 4 Brainstorming bot

A GPT bot to help with brainstorming and developing a story, acting like a consultant. This involves building an initial draft called a "straw man" and then refining it by asking questions about the audience, their background, the purpose of the document, and the main goal. The process will be repeated to improve the story, considering the content, audience, and objectives until it is fully developed.