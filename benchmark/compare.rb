#!/usr/bin/env ruby
# frozen_string_literal: true

require "benchmark"
require "json"
require "open3"
require "tmpdir"
require_relative "../lib/rbpptx"

WARMUP = Integer(ENV.fetch("RBPPTX_BENCH_WARMUP", "1"))
ITERATIONS = Integer(ENV.fetch("RBPPTX_BENCH_ITERATIONS", "5"))

DEFAULT_SAMPLE = File.expand_path(
  "../samples/prototyping_presentation.pptx", __dir__
)
SAMPLE = ENV.fetch("RBPPTX_BENCH_SAMPLE", DEFAULT_SAMPLE)

def rss_kb
  status = File.read("/proc/#{$$}/status")
  match = status.match(/^VmRSS:\s+(\d+)\s+kB$/)
  match ? match[1].to_i : 0
rescue Errno::ENOENT
  0
end

def benchmark(label)
  WARMUP.times do
    GC.start(full_mark: true, immediate_sweep: true)
    yield
  end

  samples = []
  rss_deltas = []
  ITERATIONS.times do
    GC.start(full_mark: true, immediate_sweep: true)
    started_rss = rss_kb
    real = Benchmark.realtime { yield }
    samples << real
    rss_deltas << (rss_kb - started_rss)
  end

  mean = samples.sum / samples.length
  variance = samples.sum { |s| (s - mean)**2 } / samples.length
  {
    "label" => label,
    "real" => mean,
    "real_min" => samples.min,
    "real_stddev" => Math.sqrt(variance),
    "rss_delta_kb" => rss_deltas.max,
    "iterations" => ITERATIONS
  }
end

def read_text_with_rbpptx(path)
  Rbpptx.open(path) do |pres|
    total = 0
    pres.slides.each do |slide|
      slide.shapes.each { |shape| total += shape.text.length }
    end
    total
  end
end

def fill_one_with_rbpptx(src, dst)
  Rbpptx.open(src) do |pres|
    target = pres.slides.first.shapes.find { |s| s.text != "" } || pres.slides.first.shapes.first
    target.text = "Bench Title"
    pres.save(dst)
  end
end

def fill_many_with_rbpptx(src, dst, count: 10)
  Rbpptx.open(src) do |pres|
    edited = 0
    pres.slides.each do |slide|
      break if edited >= count

      slide.shapes.each do |shape|
        runs = shape.runs
        next if runs.empty?

        runs.first.text = "fill_#{edited}"
        edited += 1
        break
      end
    end
    pres.save(dst)
  end
end

def python_pptx_results(scratch)
  script = File.expand_path("python_pptx_compare.py", __dir__)
  cmd = ["uv", "run", "--with", "python-pptx", "python3", script, SAMPLE, scratch]
  out, err, status = capture(cmd)
  unless status.success?
    warn "python-pptx benchmark failed: #{err}"
    return []
  end
  JSON.parse(out.lines.last)
rescue JSON::ParserError => e
  warn "python-pptx benchmark output not parseable: #{e.message}\n#{out}"
  []
end

def docxtemplater_results(scratch)
  script = File.expand_path("docxtemplater_compare.js", __dir__)
  bench_dir = File.expand_path(__dir__)
  unless File.directory?(File.join(bench_dir, "node_modules"))
    warn "benchmark/node_modules missing; run `cd benchmark && npm install` first"
    return []
  end
  cmd = ["node", script, SAMPLE, scratch]
  out, err, status = Open3.capture3(
    {
      "RBPPTX_BENCH_WARMUP" => WARMUP.to_s,
      "RBPPTX_BENCH_ITERATIONS" => ITERATIONS.to_s,
      "NODE_PATH" => File.join(bench_dir, "node_modules")
    },
    *cmd, chdir: bench_dir
  )
  unless status.success?
    warn "docxtemplater benchmark failed: #{err}"
    return []
  end
  JSON.parse(out.lines.last)
rescue JSON::ParserError => e
  warn "docxtemplater benchmark output not parseable: #{e.message}\n#{out}"
  []
end

def capture(cmd)
  require "open3"
  Open3.capture3({ "RBPPTX_BENCH_WARMUP" => WARMUP.to_s, "RBPPTX_BENCH_ITERATIONS" => ITERATIONS.to_s }, *cmd)
end

def render_table(rows)
  headers = ["benchmark", "real (s)", "min (s)", "stddev (s)", "rss Δ (kB)", "iters"]
  data = rows.map do |r|
    [
      r["label"],
      format("%.4f", r["real"]),
      format("%.4f", r["real_min"]),
      format("%.4f", r["real_stddev"]),
      r["rss_delta_kb"].to_s,
      r["iterations"].to_s
    ]
  end
  widths = headers.map.with_index { |h, i| [h.length, *data.map { |row| row[i].length }].max }
  fmt = widths.map { |w| "%-#{w}s" }.join(" | ")
  puts(format(fmt, *headers))
  puts(widths.map { |w| "-" * w }.join("-+-"))
  data.each { |row| puts(format(fmt, *row)) }
end

unless File.exist?(SAMPLE)
  warn "sample not found at #{SAMPLE}"
  warn "set RBPPTX_BENCH_SAMPLE or place a .pptx at the default path"
  exit 1
end

Dir.mktmpdir("rbpptx-bench") do |scratch|
  puts "sample: #{SAMPLE} (#{File.size(SAMPLE) / 1024.0 / 1024.0} MB)"
  puts "warmup: #{WARMUP}, iterations: #{ITERATIONS}"
  puts

  rb_rows = [
    benchmark("rbpptx read") { read_text_with_rbpptx(SAMPLE) },
    benchmark("rbpptx fill_one") { fill_one_with_rbpptx(SAMPLE, File.join(scratch, "rb_one.pptx")) },
    benchmark("rbpptx fill_many") { fill_many_with_rbpptx(SAMPLE, File.join(scratch, "rb_many.pptx")) }
  ]

  py_rows = python_pptx_results(scratch)
  node_rows = docxtemplater_results(scratch)
  rows = rb_rows + py_rows + node_rows

  render_table(rows)
  puts
  puts "JSON:"
  puts JSON.pretty_generate(rows)
end
