Wave Generator Ver. 1.0.0
© 2009 Larry Serflaten

The WaveGenerator ActiveX DLL is designed to help produce repeating wave patterns for multi-purpose use.

 

ACTIVEX COMPATIBILITY NOTICE

This package has been uploaded in demonstration form, its 'Version Compatibility' is set to 'No Compatibility'. That means you can't make an executable of this project as supplied.  You first must build the WaveGenerator ActiveX component
and then set a project reference to that new DLL before the demo can be run as
an exe. In other words, this demo was not intended to be ran outside of the VB development environment, it is only a demo of things that can be done using the Wave Generator component.

When you are satisfied with its functionality, start a new instance of VB and load only the Wave Generator project. Make the project ("WaveGenerator.dll") and move the DLL file to your System32 folder.  Then set binary compatibility to the DLL file.  For added safety, change all the WaveGenerator source files to Read Only.
After that, you can access the Wave Generator component by selecting it from VB's
References dialog.  If you later want to add functionality to the project, copy the files to a new location, remove the Read Only attribute and create a new project (Ex. "MyWaveGenerator.dll"). Creating a new project will help to avoid breaking any other projects that make use of the original WaveGenerator DLL.




PROJECT NOTES

Originally created to create sounds, the WaveGenerator can be used for other purposes. Its simple design helps to make it easy to use in a programming environment.  This file contains some wave basics and project notes to help you get started putting it to use.




WAVE BASICS

A wave is a changing value that follows a repeating pattern.  Waves have three basic properties; Frequency, Amplitude, and Bias. 

Frequency - The repeat rate of the pattern
Amplitude - The amount of positive and negative change in the pattern
Bias      - The pattern's offset from 0


Normally a wave will oscillate above and below 0 according to its Amplitude and Frequency.  Sometimes, however, it is desired to have the wave centered on a value other than 0.  The Bias property determines that center value. The Bias property can also be used to make an oscillator maintain a constant value.  Setting both Amplitude and Frequency to 0 (their initial state) and setting Bias to some value will cause the oscillator to constantly output that Bias value.  This biased condition will be shown to be useful later on....

Over the years, some wave patterns have been found to be useful in a wide variety of situations. These common patterns have been named according to the shape of the wave they produce: Sinusoidal, Square, Triangular, and Sawtooth:

Sinusoidal - The most common of wave shapes
Square     - A wave that alternates between two (Hi and Lo) values
Triangular - A wave whose rate of change is constant
Sawtooth   - A wave whose rate of change is linear

Amplitude, Bias, Frequency, and Shape are the properties you will use to create all the different wave forms. 



CLASS DESCRIPTIONS

The WaveGenerator project consists of six classes:

Generator.cls - Controls the changing value, common to all oscillators

SimpleOsc.cls - A basic oscillator whose basic properties (Amplitude, Frequency and Bias) are all scalar types (Single).

HybridOsc.cls - An oscillator whose basic properties are all SimpleOsc types.  This allows these properties to be modulated as the wave is generated.

ComplexOsc.cls - An oscillator whose basic properties can be any of the oscillator types. This allows modulation in more complex ways. 

I_Oscillator.cls - The polymorphic (generic) interface to access the fundamental routines to operate an oscillator.




GENERATOR DESIGN NOTES 

At the heart of any oscillator is a changing value.  Because this 'changing value' function is common to all oscillators, it was factored out into a class of its own (Generator.cls).  The generator object controls the frequency and wave shape of the oscillators.  Its input is a value (Angle) that rises from 0 to 2Pi which is the Radian equivalent of one full cycle. Its output is a value (Value) that reflects the current position in the cycle in accordance with the wave shape selected.  In summary, the generator object handles Frequency and Shape. Amplitude and Bias are handled by the individual oscillator classes.

An important factor of any (digital) generator is the number of steps it takes to complete one full cycle.  When applying the generator to sound, this step count is analogous to the (digital) sample rate of the sound. Being a multi-purpose component, 1000 was the number of steps chosen to complete one full cycle.  1000 works well when dealing with percentages of the wave (10%, 20%, 30%, ...) or wave quadrants (25%, 50%, 75%, etc), and 1000 steps also works well when applying the component to sound generation.

Typical sample rates of sounds include 11025, 22050, and 44100 samples per second. As indicated, a Frequency setting of 1 requires 1000 steps to generate one full cycle.  Creating a 1 Hertz tone requires that a buffer be filled where one full cycle takes 11025, or 22050 steps, according to the sample rate desired. To make a long story short, to fill the buffer with data, divide the desired frequency by a factor of the desired sample rate.

In short, to build a tone of 440 hertz sampled at 22050 samples per second, set the buffer length to however long you want the sound to play, and set the oscillator frequency to 440 / 220.5. That will produce a wave that has 440 full cycles in 22050 steps.  Again, that is why 1000 steps were used in the generator, because it works well in other situations and when applied to building sounds, setting the Frequency uses a fraction where the numerator is the frequency you want and the divisor closely resembles the sample rate you're using.

    

OSCILLATOR DESIGN NOTES

The Simple Oscillator (SimpleOsc) will generate a single changing value which you could very easily program yourself.  It becomes more useful as a property modulator of some other oscillator.

The Hybrid Oscillator (HybridOsc) encompasses that modulation design by declaring its basic properties (Amplitude, Frequency, and Bias) as the SimpleOsc type.  This is where the Simple Oscillator is put to use as a modulator of one or more of a wave's properties.

For example, to sweep across some number of frequencies (such as is used in making a siren) you'd want an oscillator that could vary its Frequency property, over time. Setting the hybrid's Frequency oscillator to output a sine wave would cause the hybrid's waveform to change frequency according to the value of the sine wave.

Remember, however, that waveforms typically go above and below 0.  If the sine wave goes below 0, that means the hybrid's Frequency setting becomes a negative value.  While this is not disallowed (in case you find a use for it), it may not be the desired response.  This is a case where adding a Bias to the sine wave is needed to keep its value above 0.

With an Amplitude of 500 and a Bias of 2000, the sine wave values would oscillate between 1500 and 2500, which would then be the frequency sweep output of the hybrid oscillator.

As its name implies the Complex Oscillator (ComplexOsc) is typically a little more complicated.  It can mimic the Hybrid Oscillator because its basic properties can be set to SimpleOsc types. But, it can do more than that because its basic properties also accept HybridOsc types. And to really mix things up, the ComplexOsc's basic properties also accept ComplexOsc types!



POLYMORPHIC

Polymorphic (in programming dialect) simply means a variable can hold data of different types.  In VB, where Byte, Long, and Double (to name a few) can only hold data of their specific type, the Variant data type is polymorphic in that it can be made to hold data of (nearly) any type.  When dealing with classes, VB demands that you declare a variable from that class before you can use it.  But if you want a specific variable to be able to hold one of many different types of classes, you either have to declare it as the Object type (in which case all calls are late bound) or, you have to provide a separate interface that all the classes support.

All the oscillator classes in the Wave Generator component support the I_Oscillator interface.  In most cases, to use an oscillator all you need to do is increment the value and then retrieve that new value.  The I_Oscillator interface supports both those operations as well as Clone and Initialize methods.  The Clone method returns a duplicate oscillator set to the current values, and the Initialize method presets the internal generator to specific values.




   

  
