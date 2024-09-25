from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
import sqlite3
import time


# Function to read movie names from Excel
def movies_name_from_excel():
    """
    Reads movie names from an Excel file named "Movies.xlsx".
    The Excel file should have a header named "Movie".

    Returns:
        list: A list of movie names.
    """
    excel = Files()
    excel.open_workbook("Movies.xlsx")
    data = excel.read_worksheet(name="Sheet1", header=True)
    data_list = [row["Movie"] for row in data]  # Ensure that the header of the Excel file has "Movie"
    excel.close_workbook()
    return data_list


# Class for scraping TMDb movies
class TmdbMoviesScrapper:
    def __init__(self):
        """
        Initialises the TmdbMoviesScrapper class, opens a new browser
        instance and navigates to the TMDb homepage.
        """
        self.browser = Selenium()
        self.tmdb_url = 'https://www.themoviedb.org/search?query='
        self.browser.open_available_browser()  # Opens the browser to TMDb home
        self.browser.maximize_browser_window()

    def extract_movie_data(self, movie_name):
        """
        Extracts movie details from the current page.
        Returns:
            tuple: A tuple containing the movie's TMDb score, storyline, rating, genres, and reviews.
        """
        try:
            self.search_movie(movie_name)

            # Click and extract movie details
            self.click_and_extract_movie_details(movie_name)

        except Exception as e:
            print(f"An error occurred while processing {movie_name}: {e}")



            # movie_name = self.browser.get_text('xpath://div[@class="title"]/div')
            # tmdb_score = self.browser.get_text('xpath://div[@class="user_score_chart"]')
            # storyline = self.browser.get_text('xpath://div[@class="overview"]/p')
            # rating = self.browser.get_text('xpath://span[@class="certification"]')
            # genres = ', '.join([genre.text for genre in self.browser.find_elements('xpath://div[@class="genres"]/a')])

            # # Extract top 5 critic reviews (assuming there are reviews on the page)
            # reviews = []
            # for i in range(1, 6):  # Extracting top 5 reviews if available
            #     try:
            #         review = self.browser.get_text(f"xpath://div[@class='review'][{i}]")
            #         reviews.append(review)
            #     except Exception as e:
            #         reviews.append(None)

            # return movie_name, tmdb_score, storyline, rating, genres, reviews

        # except Exception as e:
        #     print("Error while extracting movie data:", e)
        #     return None, None, None, None, None, []


    # Function to search a movie on TMDb
    def search_movie(self, movie_name):
        url = self.tmdb_url + movie_name 
        time.sleep(2) 
        self.browser.go_to(url)  # Go to the search page

        # Wait for search results to load

    def scroll_to_load_movies(self):
        """
        Scrolls down the page slowly using 'PAGE_DOWN' to load more movie results.
        """
        scroll_pause_time = 2  # Adjust the pause time as needed

    # Scroll down 5 times using 'PAGE_DOWN', with a pause between each
        for _ in range(10):  # You can increase or decrease the number of scrolls
            self.browser.press_keys(None, "PAGE_DOWN")
            time.sleep(scroll_pause_time)
            
        #     # Find all movie results by identifying the 'details' class
        #     results = self.browser.find_elements("//div[@class='card style_1']")

        #     # if not results:
        #     #     print(f"No results found for: {movie_name}")
        #     #     return False

        #     exact_match = None
        #     # latest_year = -1

        #     # Iterate through the results and find the one with the exact title and latest release date
        #     for result in results:
        #         title = self.browser.get_text(f"xpath:(//div[@class='results flex'])[1]/div//a/h2")
        #         release_year = self.browser.get_text(f"xpath:(//div[@class='results flex'])[1]/div//div/span")

        #         # Extract the year from the release date
        #         year = int(release_year.strip()[-4:])

        #         if title.strip().lower() == movie_name.strip().lower():
        #             if year > latest_year:
        #                 latest_year = year
        #                 exact_match = result

        #     if exact_match:
        #         # Click on the exact match movie
        #         self.browser.click_element("//div[@class='results flex'])[1]/div//a/h2")
        #         print(f"Found exact match: {movie_name} ({latest_year})")
        #         return True
        #     else:
        #         print(f"No exact match found for: {movie_name}")
        #         return False

        # except Exception as e:
        #     print(f"Error during movie search for {movie_name}: {e}")
        #     return False

    # Function to extract movie details

    def extract_movie_details(self, movie_name):
        """Extract TMDB score, Storyline, Genres, and Reviews."""
        # Step 5: Extract details using appropriate XPaths
        user_score_xpath = '//*[@id="consensus_pill"]/div/div[1]/div/div'
        storyline_xpath = '//div[@class="overview"]//p'
        genres_xpath = '//span[@class="genres"]'
        reviews_xpath = '//*[@id="media_v4"]/div/div/div[1]/div/section[2]/section/div[1]/ul/li'

        try:
            user_score = self.browser.get_element_attribute(user_score_xpath, "data-percent")
            storyline = self.browser.get_text(storyline_xpath)
            genres = self.browser.get_text(genres_xpath)
        

            reviews_elements = self.browser.find_elements(reviews_xpath)
        
            # Top 5 reviews
            top_reviews = [review.text for review in reviews_elements[:5]]

            # Print extracted data (or save it as needed)
            print(f"Movie: {movie_name}")
            print(f"TMDB Score: {user_score}")
            print(f"Storyline: {storyline}")
            print(f"Genres: {genres}")
            print(f"Top 5 Reviews: {top_reviews}")

        except Exception as e:
            print(f"Error while extracting details for '{movie_name}': {e}")
    
    

    def click_and_extract_movie_details(self, movie_name):
        """Click on the movie and extract its details."""
        # Locate the movie title element and click
        title_xpath = f"//div[@class='title']//a/h2"
        movie_titles_elements = self.browser.get_webelements(title_xpath)

        for index, title_element in enumerate(movie_titles_elements):
            if title_element.text == movie_name:
                print(f"Clicking on movie: {title_element.text}")
                self.browser.scroll_element_into_view(title_element)
                self.browser.click_element_when_visible(title_element)
                break

        time.sleep(5)

        # Extract movie details
        self.extract_movie_details(movie_name)

    

    def close_browser(self):
        """Closes the browser."""
        self.browser.close_browser()


        


    

@task
def scrape_movies():
    """
    Scrapes movie data from TMDb using the provided Excel file.

    This task reads movie names from an Excel file, searches for each movie on TMDb, extracts the movie's details and inserts the scraped data into a SQLite database.

    Args:
        None

    Returns:
        None
    """

    movie_names = movies_name_from_excel()

    scrapper = TmdbMoviesScrapper()

    for movie in movie_names:
        print(f"processing movie:{movie}")
        scrapper.extract_movie_data(movie)
        # Extract movie data
        

    scrapper.close_browser()

if __name__ == "__main__":
    scrape_movies()

