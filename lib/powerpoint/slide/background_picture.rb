require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class BackgroundPicture
      include Powerpoint::Util

      attr_reader :image_name, :image_path

      def initialize(options={})
        require_arguments [:presentation, :image_path], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @image_name = File.basename(@image_path)
      end

      def save(extract_path, index)
        copy_media(extract_path, @image_path)
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def file_type
        File.extname(image_name).gsub('.', '')
      end

      def save_rel_xml(extract_path, index)
        render_view('picture_background_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels")
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('picture_background_slide.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end
