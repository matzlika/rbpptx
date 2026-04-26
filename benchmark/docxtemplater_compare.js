#!/usr/bin/env node
/**
 * docxtemplater comparison driver invoked by ../benchmark/compare.rb.
 *
 * Usage:
 *   node benchmark/docxtemplater_compare.js SRC SCRATCH_DIR
 *
 * Requires `npm install` inside benchmark/ first. The Ruby driver runs
 * that for you when needed.
 */
"use strict";

const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const WARMUP = parseInt(process.env.RBPPTX_BENCH_WARMUP || "1", 10);
const ITERATIONS = parseInt(process.env.RBPPTX_BENCH_ITERATIONS || "5", 10);

function rssKb() {
  // Node reports RSS in bytes; convert to kB to match Ruby/Python sides.
  return Math.round(process.memoryUsage.rss() / 1024);
}

function benchmark(label, fn) {
  for (let i = 0; i < WARMUP; i += 1) fn();

  const samples = [];
  const rssDeltas = [];
  for (let i = 0; i < ITERATIONS; i += 1) {
    if (global.gc) global.gc();
    const before = rssKb();
    const t0 = process.hrtime.bigint();
    fn();
    const ns = Number(process.hrtime.bigint() - t0);
    samples.push(ns / 1e9);
    rssDeltas.push(rssKb() - before);
  }

  const mean = samples.reduce((a, b) => a + b, 0) / samples.length;
  const variance =
    samples.reduce((acc, s) => acc + (s - mean) ** 2, 0) / samples.length;
  return {
    label,
    real: mean,
    real_min: Math.min(...samples),
    real_stddev: Math.sqrt(variance),
    rss_delta_kb: Math.max(...rssDeltas, 0),
    iterations: ITERATIONS,
  };
}

function readWithDocxtemplater(src) {
  const buf = fs.readFileSync(src);
  const zip = new PizZip(buf);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });
  // getFullText() walks the package's text content, the closest analog to
  // walking shape.text on the rbpptx / python-pptx sides.
  return doc.getFullText().length;
}

function fillOneWithDocxtemplater(src, dst) {
  // docxtemplater is template-driven and operates on the whole package
  // in one render pass. There is no public API to "edit one shape" — the
  // honest comparison is "open the package and write it back out via
  // PizZip", which exercises the same parse + ZIP rewrite machinery the
  // template fill would. This is the cheapest thing the docxtemplater
  // stack can do, and rbpptx/python-pptx still beat it on the wider
  // benchmark, so the comparison is a fair lower bound.
  const buf = fs.readFileSync(src);
  const zip = new PizZip(buf);
  // Touch the package via Docxtemplater so we account for its parse cost.
  // eslint-disable-next-line no-new
  new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  const out = zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(dst, out);
}

function fillManyWithDocxtemplater(src, dst) {
  // Same reasoning as fillOne — docxtemplater fills templates in one pass,
  // so "many" vs "one" is the same call. Kept for table parity.
  fillOneWithDocxtemplater(src, dst);
}

function main() {
  const src = process.argv[2];
  const scratch = process.argv[3];
  fs.mkdirSync(scratch, { recursive: true });

  const results = [
    benchmark("docxtemplater read", () => readWithDocxtemplater(src)),
    benchmark("docxtemplater fill_one", () =>
      fillOneWithDocxtemplater(src, path.join(scratch, "node_one.pptx"))
    ),
    benchmark("docxtemplater fill_many", () =>
      fillManyWithDocxtemplater(src, path.join(scratch, "node_many.pptx"))
    ),
  ];
  process.stdout.write(JSON.stringify(results) + "\n");
}

main();
