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

`audioFile` is stored under the non-visual properties of the picture or shape. The object appears on the slide like normal drawing content, but actual playback is controlled from the slide timing tree. The drawing object's non-visual ID is what ties the visible placeholder to the media timing node.

The complete shape normally includes:

- `p:cNvPr`, including an `a:hlinkClick` action of `ppaction://media`,
- `p:cNvPicPr`, often with picture locks,
- `p:nvPr` with `a:audioFile`,
- a media relationship for the audio data,
- an image relationship and `a:blip` when a picture is used as the placeholder,
- shape properties such as `a:off`, `a:stretch`, and `a:fillRect`.

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
