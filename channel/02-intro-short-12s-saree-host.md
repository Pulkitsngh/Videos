# World Unscripted: 12-second host intro Short (saree character)

**✅ Final video:** https://d2ol7oe51mr4n9.cloudfront.net/user_3JzJNaJGbHkwuqjj3dJkjvMymWi/6c79df83-96fa-4485-8eaf-5380e91ddb3f.mp4
**Specs:** 720×1280 (9:16), 24 fps, H.264 + AAC, 12.5 s (12 s of speech + 0.5 s hold), −14 LUFS, Hindi captions burned in

## Host
- **Base character:** the user's Higgsfield character (Soul V2 portrait + 3-view character sheet, job `c357abb6-…`), used as the identity reference.
- **Saree look** (Nano Banana Pro, job `e2f142e1-f3f6-4a23-86c0-9c1f65c87f98`):
  - deep maroon Kanjeevaram silk saree with a gold zari border
  - red bindi, gold jhumkas, bangles, a low braid with a jasmine gajra
  - brass lantern (the channel's signature prop)
  - carved stone temple corridor at dusk
  - Reuse this image as the start frame or reference for all future host videos.
- **Voice:** ElevenLabs preset "Maya" (`b0f766b7-8703-4bd1-b973-f857c36837b6`). Reuse it for the host in every video.

## Script (as spoken)
> नमस्ते! स्वागत है वर्ल्ड अनस्क्रिप्टेड में। यहाँ मिलेंगी भारत के रहस्यमयी मंदिरों, अनसुने इतिहास और चमत्कारों की वो कहानियाँ, जो किसी स्क्रिप्ट में नहीं लिखी गईं।

## On-screen text
- **From 2.1 s, when she says the name:** a gold **WORLD UNSCRIPTED** title at the top.
- **From 10.3 s:** the tagline **दुनिया, बिना स्क्रिप्ट के।**
- **Throughout:** Hindi captions timed to her speech, with "World Unscripted" and "रहस्यमयी मंदिरों" in gold.

## How it was made (for repeating the process)
1. The first attempt used Seedance 2.0 with the narration as an audio reference. It **did not lip-sync**: the model made up its own speech, so that clip was discarded.
2. What worked (the Higgsfield narrator-workflow method):
   - Seedance 2.0 (720p, 12 s, start frame = the saree image) with **the exact Hindi line written into the prompt**, so the model speaks it with matching lips.
   - Whisper confirmed the full line was spoken in 11.9 s.
   - Then `voice_change` swapped the timbre to the "Maya" voice while keeping the lip timing.
3. Finishing:
   - Captions and titles added with the Hind Bold font.
   - Original tanpura drone mixed about 18 dB under the voice, ducking under speech.
   - Loudness normalised to −14 LUFS.

**Cost:** ≈ 4 (saree image) + 54 (first, discarded clip) + 54 (final clip) + voice change + ~1 (unused voice take).

## Upload suggestion
- **Title:** `नमस्ते! ये है World Unscripted 🙏 भारत के अनसुने रहस्य #shorts`
- **Description:** `भारत के रहस्यमयी मंदिरों, अनसुने इतिहास और चमत्कारों की वो कहानियाँ — जो किसी स्क्रिप्ट में नहीं लिखी गईं। 🪔 सब्सक्राइब करें और रहस्यों का सफ़र शुरू करें! ⚠️ इस वीडियो की होस्ट एक AI-निर्मित किरदार है।`
- **Hashtags:** `#WorldUnscripted #shorts #MysteriousIndia #IndianTemples #IndianHistory #रहस्य`
- Mark it as **Altered or synthetic content: Yes**. The host is an AI character, and YouTube requires disclosure for a realistic AI person.
- Also works as the **channel trailer** (Shorts-format) and as a pinned intro on the channel home page.
