# Add an audio file to a slide in a presentation

Audio in PresentationML is stored as media data plus slide relationships and XML markup that binds the media to a shape. The visible slide usually contains a picture or shape; the media relationship supplies the actual audio file.

## Package shape

A slide that references audio can contain markup like this:

```xml
<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="7" name="audio">
      <a:hlinkClick r:id="" action="ppaction://media"/>
    </p:cNvPr>
    <p:nvPr>
      <a:audioFile r:link="rAudio1"/>
    </p:nvPr>
  </p:nvPicPr>
</p:pic>
```

The `r:link` value points to a relationship on the slide part. The package also needs a media data part, a media reference relationship, and usually an image or shape used as the clickable placeholder.

## Rust workflow

Use `ooxmlsdk` to open the presentation and find the target slide:

```rust
{{#include ../../listings/presentation/src/lib.rs:open_presentation_read_only}}
```

`ooxmlsdk 0.6.0` exposes media data part and audio reference relationship helpers, but this chapter does not yet include a tested writer. A complete audio insertion example must verify all of these together:

- media data part content type and extension,
- slide audio and media reference relationships,
- placeholder shape or picture XML,
- timing markup for playback,
- package save and PowerPoint compatibility.

Until that fixture exists, use this chapter as the package map for implementing and testing audio insertion.
