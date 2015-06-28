package sample

import grails.test.mixin.TestFor
import spock.lang.Specification

/**
 * See the API for {@link grails.test.mixin.domain.DomainClassUnitTestMixin} for usage instructions
 */
@TestFor(Book)
class BookSpec extends Specification {

    BookExcelImporter bookImporter

    void setup() {
        bookImporter.read("./test-data/books.xls")
    }

    void testColumns() {
        setup:
        def booksMapList = bookImporter.getBooks();
        println booksMapList

        expect:
        booksMapList.every { Map bookParams -> new Book(bookParams).save() }
    }

    void testCells() {
        setup:
        def singleBook = bookImporter.getOneMoreBookParams()
        println "$singleBook"
        println "dateIssuedError: ${singleBook.dateIssuedError.class}"
        println "dateIssued: ${singleBook.dateIssued.class}"

        expect:
        singleBook.title ==  'Romeo & Juliet'
        singleBook.author == 'Shakespeare'
        singleBook.dateIssued
        singleBook.dateIssuedError
    }
}
