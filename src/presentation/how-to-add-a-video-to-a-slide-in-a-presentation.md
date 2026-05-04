# Add a video to a slide in a presentation

Video in PresentationML follows the same package pattern as audio: a slide contains markup for a clickable visual object, and relationships connect that object to media data in the package.

## Package shape

A slide can reference a video file from non-visual picture properties:

```xml
<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="7" name="video">
      <a:hlinkClick r:id="" action="ppaction://media"/>
    </p:cNvPr>
    <p:nvPr>
      <a:videoFile r:link="rVideo1"/>
    </p:nvPr>
  </p:nvPicPr>
</p:pic>
```

The slide part relationship item resolves `rVideo1` to the stored media data. A complete PowerPoint-compatible video insertion normally also includes a preview image, media reference relationship, timing data, and shape geometry.

`videoFile` is defined inside the non-visual properties of the picture or shape. The visible object sits on the slide like any other drawing object, while playback is described in the slide timing tree. The non-visual drawing ID is used by that timing data to refer back to the media object.

The video timing schema uses `CT_TLMediaNodeVideo`, whose common media node child is required and whose `fullScrn` attribute defaults to `false`.

The full insertion shape normally includes:

- `p:cNvPr`, including an `a:hlinkClick` action of `ppaction://media`,
- `p:cNvPicPr` and picture locks,
- `p:nvPr` with `a:videoFile`,
- a video relationship to the media data part,
- a media reference relationship,
- an image part plus `a:blip` for the preview frame,
- shape properties such as `a:off`, `a:stretch`, and `a:fillRect`.

## Rust workflow

Start by navigating to the target slide through the presentation part:

```rust
{{#include ../../listings/presentation/src/lib.rs:get_slide_text}}
```

`ooxmlsdk 0.6.0` has low-level package support for media data parts and reference relationships. This page does not yet publish a writer because video insertion needs a tested fixture that covers:

- video media bytes and content type,
- slide video reference relationship,
- media reference relationship,
- preview image relationship,
- `<a:videoFile/>`, picture, and timing XML,
- round-trip save validation.

Keep any implementation in `listings/` with a fixture before documenting the final API.
