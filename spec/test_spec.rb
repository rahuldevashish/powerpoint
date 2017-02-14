require 'powerpoint'

describe 'Powerpoint parsing a sample PPTX file' do
  before(:all) do 
    @deck = Powerpoint::Presentation.new
    @deck.add_intro 'Bicycle Of the Mind', 'created by Steve Jobs'
    @deck.add_textual_slide 'Why Mac?', ['Its cool!', 'Its light!']
    @deck.add_textual_slide 'Why Iphone?', ['Its fast!', 'Its cheap!']
    @deck.add_pictorial_slide 'JPG Logo', 'samples/images/sample_png.png'
    @deck.add_text_picture_slide('Text Pic Split', 'samples/images/sample_png.png', content = ['Here is a string', 'here is another'])
    @deck.add_pictorial_slide 'PNG Logo', 'samples/images/sample_png.png'
    @deck.add_picture_description_slide('Pic Desc', 'samples/images/sample_png.png', content = ['Here is a string', 'here is another'])
    @deck.add_picture_description_slide('JPG Logo', 'samples/images/sample_jpg.jpg', content = ['descriptions'])
    @deck.add_pictorial_slide 'GIF Logo', 'samples/images/sample_gif.gif', {x: 124200, y: 3356451, cx: 2895600, cy: 1013460}
    @deck.add_textual_slide 'Why Android?', ['Its great!', 'Its sweet!']
    @deck.save 'samples/pptx/sample.pptx' # Examine the PPTX file
  end

  it 'Create a PPTX file successfully.' do
    @deck.should_not be_nil
  end
end

describe 'Setting Presentation Height and Width' do

  before(:all) do
    @pts_to_emu = Powerpoint::Presentation::PTS_TO_EMU
  end

  context 'defaults' do
    before(:each) do
      @deck = Powerpoint::Presentation.new
    end

    it 'allows for a default presentation_width' do
      expect(@deck.presentation_width).to eq(Powerpoint::Presentation::DEFAULT_WIDTH * @pts_to_emu)
    end

    it 'allows for a default presentation_height' do
      expect(@deck.presentation_height).to eq(Powerpoint::Presentation::DEFAULT_HEIGHT * @pts_to_emu)
    end
  end

  context 'custom width and height' do
    before(:each) do
      @deck = Powerpoint::Presentation.new(presentation_width: 1920, presentation_height: 1200)
    end

    it 'allows the presentation_width to be set' do
      expect(@deck.presentation_width).to eq(1920 * @pts_to_emu)
    end

    it 'allows the presentation_height to be set' do
      expect(@deck.presentation_height).to eq(1200 * @pts_to_emu)
    end
  end
end