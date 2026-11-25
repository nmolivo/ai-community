<!--
  EXAMPLE ONLY: this file is not read by the build.
  Copy the parts you like into notes.md. Shows: short script, a script that has to shrink,
  lists, inline code, fenced code (python / bash / json), and a very short script
  (the rest of the speaker column becomes ruled notes lines).
-->

# 1 | Stable Diffusion in Action

## Handout

Welcome. This handout follows the slides one page at a time; use the space under each slide for your own notes.

## Script

Welcome, everyone. Today we connect NLP, CNNs and Transformers to image generation, and finish with a fine-tuning technique you can run yourself.


# 3 | This presentation comes with a demo

## Handout

**Before the demo you need:**

- a Hugging Face token (free)
- a GPU runtime, for example the free tier of Google Colab
- 8 to 20 images of your own subject

## Script

Everything we cover today is in the accompanying notebook, so you can follow along or run it afterwards.

Take a moment now to check that you can log in to Hugging Face. If you do not have a token yet, the sign-up is free and takes about a minute. I will pause here and wait for hands.

The notebook needs a GPU. The free Colab tier works for this demo, and it takes roughly an hour end to end. If you would rather run it on your own provider, a pay-as-you-go instance costs about ten euros for the whole exercise.

Finally, please use your own pictures. Eight to twenty is enough, and a consistent subject works best. Cats, mugs and houseplants have all worked well before.

If someone asks why we do not just train a whole new model: remind them that textual inversion changes one embedding, not billions of weights. We come back to that in the second half.


# 10 | Training loop steps

## Handout

The training loop has five steps. The learning rate only acts at step 4.

1. Forward pass
2. Loss
3. Backward pass (gradients)
4. **Update** the weights: `W <- W - lr * grad`
5. (Optional) a scheduler changes the learning rate

```python
for batch in loader:
    pred = model(batch.x)            # 1. forward
    loss = loss_fn(pred, batch.y)    # 2. loss
    optimizer.zero_grad()
    loss.backward()                  # 3. gradients
    optimizer.step()                 # 4. update (lr acts here)
    scheduler.step()                 # 5. optional
```

## Script

Walk through the loop top to bottom. Point at line four: this is the only place the learning rate acts.

Ask the room: what happens if the step is too large? Take two answers, then show the overshoot picture.

**Pause** here for questions before moving to text-to-image training.


# 14 | Fine Tuning: Textual Inversion

## Handout

Textual inversion learns **one new token embedding** and leaves the model weights frozen.

```python
pipe.load_textual_inversion("sd-concepts-library/cat-toy")
image = pipe("a <cat-toy> on a beach towel").images[0]
```

```json
{"placeholder_token": "<cat-toy>", "initializer_token": "toy", "num_vectors": 1}
```

```bash
pip install diffusers transformers accelerate
huggingface-cli login
```

## Script

A deliberately short script: the rest of the column becomes a ruled notes area.
